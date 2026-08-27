import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Flask serves this directory from course_selection/workbench_static in the
// installed application. Keep the output stable so `npm run build` is the
// only release step needed to update the local workbench UI.
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "../course_selection/workbench_static",
    emptyOutDir: true,
    sourcemap: false,
  },
  server: {
    proxy: { "/api": "http://127.0.0.1:5000" },
  },
});
