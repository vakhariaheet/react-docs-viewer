import { defineConfig } from 'vite'
import path from 'path'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [ react() ],
  build: {
    lib: {
      entry: path.resolve(__dirname, 'src/lib/index.ts'),
      name: 'react-spreadsheet-renderer',
      fileName: (format) => `react-spreadsheet-renderer.${format}.js`
    },
    rollupOptions: {
      external: [ 'react', 'react-dom' ],
      output: {
        globals: {
          react: 'React',
        }
      }
    }
  }
})
