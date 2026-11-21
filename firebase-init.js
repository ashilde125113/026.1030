import { getApp, getApps, initializeApp } from "https://www.gstatic.com/firebasejs/13.4.0/firebase-app.js";
import {
  getAuth,
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  signOut,
} from "https://www.gstatic.com/firebasejs/13.4.0/firebase-auth.js";

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

const auth = getAuth(firebaseApp);
const googleProvider = new GoogleAuthProvider();
googleProvider.setCustomParameters({ prompt: "select_account" });

const signInButton = document.querySelector("#googleSignInButton");
const signOutButton = document.querySelector("#signOutButton");
const userName = document.querySelector("#authUserName");
const avatar = document.querySelector("#authAvatar");
const status = document.querySelector("#authStatus");

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

onAuthStateChanged(auth, (user) => {
  signInButton.hidden = Boolean(user);
  signOutButton.hidden = !user;
  userName.hidden = !user;
  avatar.textContent = user ? (user.displayName || user.email || "使").slice(0, 1) : "安";
  avatar.title = user ? (user.displayName || user.email || "已登入") : "尚未登入";
  avatar.setAttribute("aria-label", avatar.title);
  if (user) userName.textContent = user.displayName || user.email || "已登入";
});
