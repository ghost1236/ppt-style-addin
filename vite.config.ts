import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';

// HTTPS 인증서 (office-addin-dev-certs 설치 후 생성됨)
function getHttpsConfig() {
  const certPath = path.join(process.env.HOME || '', '.office-addin-dev-certs');
  const keyFile = path.join(certPath, 'localhost.key');
  const certFile = path.join(certPath, 'localhost.crt');
  if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
    return { key: fs.readFileSync(keyFile), cert: fs.readFileSync(certFile) };
  }
  return true; // vite 자체 자가서명 인증서 사용
}

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    https: getHttpsConfig(),
  },
  build: {
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, 'src/taskpane/index.html'),
        commands: path.resolve(__dirname, 'src/commands/commands.html'),
      },
    },
    outDir: 'dist',
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
});
