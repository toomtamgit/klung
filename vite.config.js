import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" ทำให้ใช้ได้ทั้ง Cloudflare Pages และ GitHub Pages (ที่อยู่ในโฟลเดอร์ย่อย /ชื่อ-repo/)
export default defineConfig({
  plugins: [react()],
  base: "./",
});
