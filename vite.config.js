import { defineConfig } from 'vite';
import tsconfigPaths from 'vite-tsconfig-paths';

export default defineConfig({
  plugins: [tsconfigPaths()],
  test: {
    globals: true, // Enables `describe`, `it`, etc. globally
    environment: 'node', // Sets test environment to Node.js, adjust to 'jsdom' if needed
  },
});