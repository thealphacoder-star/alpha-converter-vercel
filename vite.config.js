import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('/node_modules/pdfjs-dist/')) return 'pdfjs';
          if (id.includes('/node_modules/html2canvas/')) return 'html-canvas';
          if (id.includes('/node_modules/canvg/')) return 'svg-renderer';
          if (id.includes('/node_modules/dompurify/')) return 'sanitizer';
          if (id.includes('/node_modules/jspdf-autotable/')) return 'pdf-tables';
          if (id.includes('/node_modules/jspdf/')) return 'pdf-generator';
          if (id.includes('/node_modules/xlsx/')) return 'spreadsheet';
          if (id.includes('/node_modules/jszip/')) return 'archive';
        },
      },
    },
  },
});
