import { dataTableComponentSource } from './table-generator.js';

export interface ScaffoldFile {
  relativePath: string;
  content: string;
}

export function scaffoldStaticFiles(appTitle: string): ScaffoldFile[] {
  return [
    { relativePath: 'package.json',                    content: packageJson(appTitle) },
    { relativePath: 'vite.config.ts',                  content: viteConfig() },
    { relativePath: 'tsconfig.json',                   content: tsConfig() },
    { relativePath: 'api/tsconfig.json',               content: apiTsConfig() },
    { relativePath: 'index.html',                      content: indexHtml(appTitle) },
    { relativePath: 'src/main.tsx',                    content: mainTsx() },
    { relativePath: 'src/components/DataTable.tsx',    content: dataTableComponentSource() },
  ];
}

// ─── Static file templates ────────────────────────────────────────────────────

function packageJson(title: string): string {
  const name = title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  return JSON.stringify({
    name,
    version: '0.0.1',
    private: true,
    type: 'module',
    scripts: {
      dev:     'concurrently "vite" "tsx api/server.ts"',
      build:   'tsc -b && vite build',
      preview: 'vite preview',
      api:     'tsx api/server.ts',
    },
    dependencies: {
      react:              '^19.0.0',
      'react-dom':        '^19.0.0',
      'react-router-dom': '^7.0.0',
      express:            '^4.18.0',
      cors:               '^2.8.5',
      duckdb:             '^1.1.0',
    },
    devDependencies: {
      '@types/react':         '^19.0.0',
      '@types/react-dom':     '^19.0.0',
      '@types/express':       '^4.17.0',
      '@types/cors':          '^2.8.0',
      '@types/node':          '^22.0.0',
      '@vitejs/plugin-react': '^4.3.0',
      typescript:             '^5.7.0',
      vite:                   '^6.0.0',
      concurrently:           '^8.0.0',
      tsx:                    '^4.19.0',
    },
  }, null, 2);
}

function viteConfig(): string {
  return `import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/api': {
        target: 'http://localhost:3001',
        changeOrigin: true,
      },
    },
  },
})
`;
}

function apiTsConfig(): string {
  return JSON.stringify({
    compilerOptions: {
      target: 'ES2022',
      module: 'NodeNext',
      moduleResolution: 'NodeNext',
      outDir: '../dist-api',
      strict: true,
      skipLibCheck: true,
      esModuleInterop: true,
    },
    include: ['.'],
  }, null, 2);
}

function tsConfig(): string {
  return JSON.stringify({
    compilerOptions: {
      target: 'ES2020',
      useDefineForClassFields: true,
      lib: ['ES2020', 'DOM', 'DOM.Iterable'],
      module: 'ESNext',
      skipLibCheck: true,
      moduleResolution: 'bundler',
      allowImportingTsExtensions: true,
      isolatedModules: true,
      moduleDetection: 'force',
      noEmit: true,
      jsx: 'react-jsx',
      strict: true,
    },
    include: ['src'],
  }, null, 2);
}

function indexHtml(title: string): string {
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${escHtml(title)}</title>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
      *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: 'Poppins', sans-serif; background: #f8fafc; color: #0f172a; }
    </style>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
`;
}

function mainTsx(): string {
  return `import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import { App } from './App'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <BrowserRouter>
      <App />
    </BrowserRouter>
  </StrictMode>
)
`;
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
