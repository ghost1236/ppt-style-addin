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
  return true;
}

export default defineConfig({
  plugins: [react()],
  // GitHub Pages 배포 base 경로
  base: '/ppt-style-addin/',
  server: {
    port: 3000,
    https: getHttpsConfig(),
  },
  build: {
    rollupOptions: {
      input: {
        // 빌드 입력: 루트 레벨 HTML → dist/taskpane/, dist/commands/ 구조 출력
        taskpane: path.resolve(__dirname, 'taskpane/index.html'),
        commands: path.resolve(__dirname, 'commands/commands.html'),
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
