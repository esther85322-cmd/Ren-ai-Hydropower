import React from "react";
import ReactDOM from "react-dom/client";

const style = document.createElement("style");
style.textContent = `
  * { box-sizing: border-box; }
  html, body, #root { height: 100%; margin: 0; padding: 0; background: #1C2530; }
  body { -webkit-tap-highlight-color: transparent; overscroll-behavior-y: contain; }
`;
document.head.appendChild(style);

// 簡化前台頁面：/checkin（打卡）、/tasks（待確認事項）——給不需要用到完整後台的人
// （例如只負責記出工的師傅/家人）。用 dynamic import 讓這些路徑只下載自己的程式碼，
// 不會把整個管理系統（含 Excel/Word 匯出、圖表等大型套件）一起下載下來，開啟速度快很多。
const path = window.location.pathname.replace(/\/+$/, "");
const page = path.endsWith("/checkin")
  ? import("./CheckinPage.jsx")
  : path.endsWith("/tasks")
  ? import("./TaskPage.jsx")
  : import("./App.jsx");
const root = ReactDOM.createRoot(document.getElementById("root"));

page.then(({ default: Page }) => {
  root.render(
    <React.StrictMode>
      <Page />
    </React.StrictMode>
  );
});
