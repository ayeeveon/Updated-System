import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react-swc'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // Proxy API requests to Flask backend
      '/login': {
        target: 'http://localhost:5000',
        changeOrigin: true
      },
      '/upload': {
        target: 'http://localhost:5000',
        changeOrigin: true
      },
      '/generate': {
        target: 'http://localhost:5000',
        changeOrigin: true
      },
      '/forgot_password': {
        target: 'http://localhost:5000',
        changeOrigin: true
      },
      '/verify_code': {
        target: 'http://localhost:5000',
        changeOrigin: true
      },
      '/send_email': {
        target: 'http://localhost:5000',
        changeOrigin: true
      }
    }
  }
})
