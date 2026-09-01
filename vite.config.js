import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// 如果你的 GitHub repo 名稱不是 "water-electric-ledger"，
// 請把下面 base 的值改成 "/你的repo名稱/"（前後都要有斜線）。
// 如果你是用 Firebase Hosting（而不是 GitHub Pages）部署，base 改回 "/" 即可。
export default defineConfig({
  plugins: [react()],
  base: "/",
});
