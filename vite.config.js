import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('/node_modules/pdfjs-dist/')) return 'pdfjs';
          if (
            id.includes('/node_modules/jspdf/') ||
            id.includes('/node_modules/jspdf-autotable/') ||
            id.includes('/node_modules/html2canvas/') ||
            id.includes('/node_modules/canvg/') ||
            id.includes('/node_modules/dompurify/')
          ) return 'pdf-tools';
          if (id.includes('/node_modules/xlsx/')) return 'spreadsheet';
          if (id.includes('/node_modules/jszip/')) return 'archive';
        },
      },
    },
  },
});
