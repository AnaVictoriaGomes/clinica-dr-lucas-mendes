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
const builtHtml = rawHtml.replace(
  /<script\s+type="module"\s+src="\/src\/main\.tsx"><\/script>/,
  '<script src="/assets/kommo-bundle.js"></script>'
);

if (!builtHtml.includes('/assets/kommo-bundle.js')) {
  throw new Error('Falha ao substituir a referência de /src/main.tsx no index.html');
}

fs.writeFileSync(path.join(distDir, 'index.html'), builtHtml, 'utf8');

console.log('Build estático (esbuild) concluído com sucesso em dist/');
