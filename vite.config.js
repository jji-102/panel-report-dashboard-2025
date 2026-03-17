import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './', // <-- ใส่ชื่อ Repo ของคุณตรงนี้ (ต้องมี / ปิดหน้าหลัง)
})