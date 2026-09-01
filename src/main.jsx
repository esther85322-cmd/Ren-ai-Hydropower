import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import CheckinPage from "./CheckinPage.jsx";

const style = document.createElement("style");
style.textContent = `
  * { box-sizing: border-box; }
  html, body, #root { height: 100%; margin: 0; padding: 0; background: #1C2530; }
  body { -webkit-tap-highlight-color: transparent; overscroll-behavior-y: contain; }
`;
document.head.appendChild(style);

// 簡化打卡頁面：/checkin — 給不需要用到完整後台的人（例如只負責記出工的師傅/家人）
const isCheckin = window.location.pathname.replace(/\/+$/, "").endsWith("/checkin");

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    {isCheckin ? <CheckinPage /> : <App />}
  </React.StrictMode>
);
