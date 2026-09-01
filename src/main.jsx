import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

const style = document.createElement("style");
style.textContent = `
  * { box-sizing: border-box; }
  html, body, #root { height: 100%; margin: 0; padding: 0; background: #1C2530; }
  body { -webkit-tap-highlight-color: transparent; overscroll-behavior-y: contain; }
`;
document.head.appendChild(style);

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
