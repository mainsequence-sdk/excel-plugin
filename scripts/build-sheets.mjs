import { build } from 'esbuild';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load BASE_ROOT from repo-level .env
dotenv.config({ path: path.resolve(__dirname, '..', '.env') });
const BASE_ROOT = process.env.BASE_ROOT || '';

const proj = path.resolve(__dirname, '..', 'sheets-addon');
const outDir = path.join(proj, 'dist');
await fs.promises.rm(outDir, { recursive: true, force: true });
await fs.promises.mkdir(outDir, { recursive: true });

// Define top-level functions so Google Sheets recognizes them.
// They call into the bundled module (global name MSCF).
const footer = `
// ---- Apps Script entry points (TOP-LEVEL) ----

/** @customfunction */
function SET_BASE_ROOT(url) {
  return MSCF.__SET_BASE_ROOT(url);
}

/** @customfunction */
function GET_DATA_BETWEEN_DATES_FROM_API(
  table_name, start_date, end_date, great_or_equal, less_or_equal, columns, unique_identifier_list
) {
  return MSCF.__GET_DATA_BETWEEN_DATES_FROM_API(
    table_name, start_date, end_date, great_or_equal, less_or_equal, columns, unique_identifier_list
  );
}

// Optional debug helpers (menu + log/ping)
function onOpen()            { return MSCF.__onOpen(); }
function DEBUG_LOG_BASEROOT() { return MSCF.__DEBUG_LOG_BASEROOT(); }
function DEBUG_PING()         { return MSCF.__DEBUG_PING(); }
`;

await build({
  entryPoints: [path.join(proj, 'src', 'index.ts')],
  bundle: true,
  platform: 'browser',
  target: 'es2019',
  format: 'iife',
  globalName: 'MSCF',
  outfile: path.join(outDir, 'Code.js'),
  define: {
    'process.env.BASE_ROOT': JSON.stringify(BASE_ROOT)
  },
  footer: { js: footer }
});

// Copy manifest alongside bundle
await fs.promises.copyFile(
  path.join(proj, 'appsscript.json'),
  path.join(outDir, 'appsscript.json')
);

console.log('✅ Built Google Sheets to sheets-addon/dist/Code.js with BASE_ROOT:', BASE_ROOT || '(empty)');
