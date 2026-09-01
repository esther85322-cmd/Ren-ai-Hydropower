// Firebase 初始化設定
//
// 請把下面 firebaseConfig 裡的值換成你自己 Firebase 專案的設定值。
// 取得方式：Firebase Console → 專案設定（齒輪圖示）→ 一般 → 你的應用程式 → SDK 設定與程式碼片段 → 選「設定」
//
// 這個檔案本身可以安全地提交到 GitHub（公開的 repo 也沒關係）：
// 這些 apiKey 等設定值本來就是「給瀏覽器用的」，不是機密金鑰。
// 真正的資料保護是靠 Firestore 的「安全規則」（見 README.md），不是靠隱藏這些設定值。

const firebaseConfig = {
  apiKey: "AIzaSyCojiUdfYtatc9mUVrHzKP_SAkz6XccWMw",
  authDomain: "ren-ai-hydropower.firebaseapp.com",
  databaseURL: "https://ren-ai-hydropower-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "ren-ai-hydropower",
  storageBucket: "ren-ai-hydropower.firebasestorage.app",
  messagingSenderId: "777332429023",
  appId: "1:777332429023:web:0b499528faf10af59d196a",
  measurementId: "G-86LCBBDEJR",
};

import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth, signInAnonymously, onAuthStateChanged } from "firebase/auth";

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);

// 用「匿名登入」讓 Firestore 安全規則可以要求 request.auth != null，
// 同時使用者完全不需要自己輸入帳號密碼——第一次打開網頁時會自動、無聲地登入。
let signInPromise = null;
export function ensureSignedIn() {
  if (auth.currentUser) return Promise.resolve(auth.currentUser);
  if (!signInPromise) {
    signInPromise = new Promise((resolve, reject) => {
      const unsub = onAuthStateChanged(
        auth,
        (user) => {
          if (user) {
            unsub();
            resolve(user);
          }
        },
        reject
      );
      signInAnonymously(auth).catch((err) => {
        unsub();
        reject(err);
      });
    });
  }
  return signInPromise;
}
