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
import {
  initializeFirestore,
  persistentLocalCache,
  persistentMultipleTabManager,
} from "firebase/firestore";
import { getAuth, signInAnonymously, onAuthStateChanged } from "firebase/auth";

const app = initializeApp(firebaseConfig);

// 本機快取：第一次打開會照常連線抓資料，但抓過一次之後會存在瀏覽器裡，
// 之後再打開網頁時會先秒開快取畫面，同時在背景悄悄跟雲端同步最新資料，
// 不用每次都整頁乾等 Firestore 回應。
// experimentalForceLongPolling：跳過 Firestore 預設的連線方式自動偵測
// （這個偵測在某些行動網路 / Wi-Fi 環境下會拖到 30-60 秒才 fallback），
// 直接一開始就用相容性最好的長輪詢方式連線，明顯縮短第一次打開的等待時間。
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
  experimentalForceLongPolling: true,
});
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
