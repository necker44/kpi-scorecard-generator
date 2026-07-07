import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base so the build works whether it's served from a custom domain
// or from https://<user>.github.io/<repo>/
export default defineConfig({
  plugins: [react()],
  base: "./",
});
