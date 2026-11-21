import { getApp, getApps, initializeApp } from "https://www.gstatic.com/firebasejs/13.4.0/firebase-app.js";

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
