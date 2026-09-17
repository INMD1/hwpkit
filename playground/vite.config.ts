import { defineConfig } from 'vite';
import { resolve } from 'path';
import { createLibreOfficeDocConverter } from '../src/node';

const convertDoc = createLibreOfficeDocConverter();

export default defineConfig({
  root: resolve(__dirname),
  publicDir: resolve(__dirname, 'public'),
  plugins: [{
    name: 'local-legacy-doc',
    configureServer(server) {
      server.middlewares.use('/api/doc-to-docx', async (req, res) => {
        if (req.method !== 'POST') { res.statusCode = 405; res.end('POST required'); return; }
        try {
          // Reject requests from pages on other origins.
          if (req.headers.origin && new URL(req.headers.origin).host !== req.headers.host) {
            res.statusCode = 403; res.end('Origin mismatch'); return;
          }
          const chunks: Buffer[] = [];
          let size = 0;
          for await (const chunk of req) {
            size += chunk.length;
            if (size > 25 * 1024 * 1024) { res.statusCode = 413; res.end('DOC 파일은 25 MiB 이하로 올려주세요.'); return; }
            chunks.push(Buffer.from(chunk));
          }
          const result = await convertDoc(Buffer.concat(chunks));
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          res.end(result);
        } catch (error) {
          res.statusCode = 422;
          res.end(error instanceof Error ? error.message : String(error));
        }
      });
    },
  }],
  resolve: {
    alias: {
      'hwpkit': resolve(__dirname, '../src/index.ts'),
    },
  },
  define: {
    'process.env': {},
    global: 'globalThis',
  },
});
