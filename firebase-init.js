import { getApp, getApps, initializeApp } from "https://www.gstatic.com/firebasejs/13.4.0/firebase-app.js";
import {
  getAuth,
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  signOut,
} from "https://www.gstatic.com/firebasejs/13.4.0/firebase-auth.js";
import {
  collection,
  getDocs,
  getFirestore,
  onSnapshot,
  writeBatch,
  doc,
} from "https://www.gstatic.com/firebasejs/13.4.0/firebase-firestore.js";
import { getStorage, ref, uploadBytes, getDownloadURL } from "https://www.gstatic.com/firebasejs/13.4.0/firebase-storage.js";
import { updateProfile } from "https://www.gstatic.com/firebasejs/13.4.0/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyBZyPfjbLWvKJTc7EE34yKZEwWLPUJuHPs",
  authDomain: "ashilde.firebaseapp.com",
  projectId: "ashilde",
  storageBucket: "ashilde.firebasestorage.app",
  messagingSenderId: "1000996421842",
  appId: "1:1000996421842:web:859d7f66c56303d45e2d0b",
};

export const firebaseApp = getApps().length ? getApp() : initializeApp(firebaseConfig);
window.firebaseApp = firebaseApp;
const db = getFirestore(firebaseApp, "planner-sync");
const storage = getStorage(firebaseApp);

const auth = getAuth(firebaseApp);
const googleProvider = new GoogleAuthProvider();
googleProvider.setCustomParameters({ prompt: "select_account" });

const signInButton = document.querySelector("#googleSignInButton");
const signOutButton = document.querySelector("#signOutButton");
const userName = document.querySelector("#authUserName");
const avatar = document.querySelector("#authAvatar");
const status = document.querySelector("#authStatus");
const avatarFile = document.querySelector("#avatarFile");

function showStatus(message = "") {
  status.textContent = message;
  status.hidden = !message;
}

function authErrorMessage(error) {
  if (error.code === "auth/unauthorized-domain") return "此網站網域尚未加入 Firebase 授權清單。";
  if (error.code === "auth/popup-closed-by-user") return "登入視窗已關閉，尚未登入。";
  if (error.code === "auth/popup-blocked") return "瀏覽器封鎖登入視窗，請允許此網站開啟彈出視窗。";
  return "Google 登入失敗，請稍後再試。";
}

let activeUid = null;
let syncGeneration = 0;
let unsubscribers = [];
let syncTimer;
let syncReady = false;
let pendingState = null;

function stopSync() {
  syncReady = false;
  unsubscribers.forEach((unsubscribe) => unsubscribe());
  unsubscribers = [];
  clearTimeout(syncTimer);
  pendingState = null;
}

function reportSyncError(error) {
  console.error("Planner sync failed", error);
  showStatus("同步失敗，請確認網路或稍後重試。");
}

function itemCollection(uid) {
  return collection(db, "users", uid, "items");
}

function habitCollection(uid) {
  return collection(db, "users", uid, "habits");
}

function habitRecords(doneMap) {
  return Object.entries(doneMap || {}).map(([id, done]) => {
    const match = id.match(/^(.*)-(\d{4}-\d{2}-\d{2})$/);
    return match ? { id, name: match[1], date: match[2], done: Boolean(done) } : null;
  }).filter(Boolean);
}

async function saveState(uid, state) {
  if (!syncReady || uid !== activeUid) return;
  try {
    const batch = writeBatch(db);
    let writes = 0;
    for (const item of state.items || []) {
      if (!Number.isInteger(item.id)) continue;
      batch.set(doc(db, "users", uid, "items", String(item.id)), item);
      writes++;
    }
    for (const habit of habitRecords(state.habitDone)) {
      batch.set(doc(db, "users", uid, "habits", habit.id), {
        name: habit.name, date: habit.date, done: habit.done,
      });
      writes++;
    }
    if (writes) await batch.commit();
    showStatus("已同步至雲端");
    setTimeout(() => { if (activeUid === uid) showStatus(); }, 1800);
  } catch (error) {
    reportSyncError(error);
  }
}

window.addEventListener("planner:state-changed", (event) => {
  if (!event.detail?.uid || event.detail.uid !== activeUid) return;
  if (!syncReady) { pendingState = event.detail; return; }
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => saveState(activeUid, event.detail), 250);
});

async function startSync(user) {
  stopSync();
  const uid = user.uid;
  activeUid = uid;
  const generation = ++syncGeneration;
  if (!window.plannerStore) {
    showStatus("同步初始化中…");
    setTimeout(() => { if (activeUid === uid) startSync(user); }, 0);
    return;
  }
  window.plannerStore.beginUser(uid);
  showStatus("正在載入雲端資料…");

  try {
    const [remoteItemsSnap, remoteHabitsSnap] = await Promise.all([
      getDocs(itemCollection(uid)), getDocs(habitCollection(uid)),
    ]);
    if (generation !== syncGeneration || activeUid !== uid) return;

    const itemIds = new Set(remoteItemsSnap.docs.map((entry) => entry.id));
    const habitIds = new Set(remoteHabitsSnap.docs.map((entry) => entry.id));
    const migration = window.plannerStore.getMigrationPayload(uid);
    if (migration) {
      const batch = writeBatch(db);
      let writes = 0;
      for (const item of migration.items) {
        const id = String(item.id);
        if (Number.isInteger(item.id) && !itemIds.has(id)) {
          batch.set(doc(db, "users", uid, "items", id), item);
          itemIds.add(id); writes++;
        }
      }
      for (const habit of habitRecords(migration.habitDone)) {
        if (!habitIds.has(habit.id)) {
          batch.set(doc(db, "users", uid, "habits", habit.id), {
            name: habit.name, date: habit.date, done: habit.done,
          });
          habitIds.add(habit.id); writes++;
        }
      }
      if (writes) await batch.commit();
      if (generation !== syncGeneration || activeUid !== uid) return;
      window.plannerStore.markMigrationComplete(uid);
    }

    const combinedItems = new Map(remoteItemsSnap.docs.map((entry) => [entry.id, { ...entry.data(), id: Number(entry.id) }]));
    const combinedHabits = Object.fromEntries(remoteHabitsSnap.docs.map((entry) => {
      const value = entry.data(); return [`${value.name}-${value.date}`, value.done];
    }));
    for (const item of migration?.items || []) if (!combinedItems.has(String(item.id))) combinedItems.set(String(item.id), item);
    for (const habit of habitRecords(migration?.habitDone)) if (!(habit.id in combinedHabits)) combinedHabits[habit.id] = habit.done;
    window.plannerStore.applyRemote(uid, [...combinedItems.values()], combinedHabits);
    syncReady = true;
    const queuedState = pendingState;
    pendingState = null;
    if (queuedState) saveState(uid, queuedState);

    let itemsReady = false;
    let habitsReady = false;
    let liveItems = [];
    let liveHabits = {};
    const applyLive = () => {
      if (!itemsReady || !habitsReady || generation !== syncGeneration || activeUid !== uid) return;
      window.plannerStore.applyRemote(uid, liveItems, liveHabits);
    };
    unsubscribers = [
      onSnapshot(itemCollection(uid), (snapshot) => {
        liveItems = snapshot.docs.map((entry) => ({ ...entry.data(), id: Number(entry.id) }));
        itemsReady = true; applyLive();
      }, reportSyncError),
      onSnapshot(habitCollection(uid), (snapshot) => {
        liveHabits = Object.fromEntries(snapshot.docs.map((entry) => {
          const value = entry.data(); return [`${value.name}-${value.date}`, value.done];
        }));
        habitsReady = true; applyLive();
      }, reportSyncError),
    ];
    showStatus("已連線，資料會自動同步");
    setTimeout(() => { if (activeUid === uid) showStatus(); }, 2500);
  } catch (error) {
    if (generation === syncGeneration) reportSyncError(error);
  }
}

signInButton.addEventListener("click", async () => {
  signInButton.disabled = true;
  showStatus("正在開啟 Google 登入…");
  try {
    await signInWithPopup(auth, googleProvider);
    showStatus();
  } catch (error) {
    showStatus(authErrorMessage(error));
  } finally {
    signInButton.disabled = false;
  }
});

signOutButton.addEventListener("click", async () => {
  try {
    await signOut(auth);
    showStatus();
  } catch {
    showStatus("登出失敗，請稍後再試。 ");
  }
});

avatar.addEventListener("click", () => {
  if (auth.currentUser) avatarFile.click();
  else showStatus("請先登入，再設定頭像照片。");
});

avatarFile.addEventListener("change", async () => {
  const file = avatarFile.files?.[0];
  const user = auth.currentUser;
  avatarFile.value = "";
  if (!file || !user) return;
  if (!file.type.startsWith("image/")) {
    showStatus("請選擇圖片檔案。");
    return;
  }
  if (file.size > 5 * 1024 * 1024) {
    showStatus("照片大小需在 5 MB 以內。");
    return;
  }

  avatar.disabled = true;
  showStatus("正在更新頭像…");
  try {
    const imageRef = ref(storage, `users/${user.uid}/avatar`);
    await uploadBytes(imageRef, file, { contentType: file.type });
    const photoURL = await getDownloadURL(imageRef);
    await updateProfile(user, { photoURL });
    renderAvatar(user);
    showStatus("頭像已更新");
    setTimeout(() => { if (auth.currentUser?.uid === user.uid) showStatus(); }, 1800);
  } catch (error) {
    console.error("Avatar upload failed", error);
    showStatus("頭像更新失敗，請確認 Firebase Storage 已啟用後重試。");
  } finally {
    avatar.disabled = false;
  }
});

function renderAvatar(user) {
  avatar.replaceChildren();
  if (user?.photoURL) {
    const image = document.createElement("img");
    image.src = user.photoURL;
    image.alt = "個人頭像";
    avatar.append(image);
  } else {
    avatar.textContent = user ? (user.displayName || user.email || "使").slice(0, 1) : "安";
  }
  avatar.title = user ? "點擊更換頭像" : "尚未登入";
  avatar.setAttribute("aria-label", avatar.title);
}

onAuthStateChanged(auth, (user) => {
  stopSync();
  syncGeneration++;
  activeUid = user?.uid || null;
  signInButton.hidden = Boolean(user);
  signOutButton.hidden = !user;
  userName.hidden = !user;
  renderAvatar(user);
  if (user) userName.textContent = user.displayName || user.email || "已登入";
  if (user) startSync(user);
  else {
    window.plannerStore?.endUser();
    showStatus();
  }
});
