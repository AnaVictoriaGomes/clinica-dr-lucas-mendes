import { build } from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';

const rootDir = process.cwd();
const distDir = path.join(rootDir, 'dist');
const assetsDir = path.join(distDir, 'assets');

const noMinify = process.argv.includes('false');

fs.rmSync(distDir, { recursive: true, force: true });
fs.mkdirSync(assetsDir, { recursive: true });

await build({
  entryPoints: [path.join(rootDir, 'src', 'main.tsx')],
  bundle: true,
  outfile: path.join(assetsDir, 'kommo-bundle.js'),
  format: 'iife',
  platform: 'browser',
  target: ['es2020'],
  minify: !noMinify,
  sourcemap: false,
  logLevel: 'info',
});

const rawHtml = fs.readFileSync(path.join(rootDir, 'index.html'), 'utf8');
const mainScriptPattern = /<script\b[^>]*\bsrc=["']\.?\/src\/main\.tsx["'][^>]*>\s*<\/script>/i;
let builtHtml = rawHtml;

if (mainScriptPattern.test(rawHtml)) {
  builtHtml = rawHtml.replace(
    mainScriptPattern,
    '<script src="/assets/kommo-bundle.js"></script>'
  );
} else if (!rawHtml.includes('/assets/kommo-bundle.js')) {
  console.warn('Aviso: referência de /src/main.tsx não encontrada no index.html; injetando /assets/kommo-bundle.js antes de </body>.');
  if (/<\/body>/i.test(rawHtml)) {
    builtHtml = rawHtml.replace(
      /<\/body>/i,
      '<script src="/assets/kommo-bundle.js"></script>\n</body>'
    );
  } else {
    builtHtml = `${rawHtml}\n<script src="/assets/kommo-bundle.js"></script>\n`;
  }
}

if (!builtHtml.includes('/assets/kommo-bundle.js')) {
  throw new Error('Falha ao injetar /assets/kommo-bundle.js no index.html');
}

fs.writeFileSync(path.join(distDir, 'index.html'), builtHtml, 'utf8');

console.log('Build estático (esbuild) concluído com sucesso em dist/');

