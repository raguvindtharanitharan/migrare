import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, mkdtemp, readdir, rm, writeFile as writeFileAsync } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import AdmZip from 'adm-zip';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// Script lives at src/extractors/ — one level up from dist/ when compiled
const PYTHON_SCRIPT = path.join(__dirname, '..', 'src', 'extractors', 'hyper-to-parquet.py');

export interface ParquetTable {
  /** Fully-qualified Hyper table name: "schema.table" */
  qualifiedName: string;
  /** Absolute path to the written .parquet file */
  parquetPath: string;
  /** Slug derived from table name — used as filename and URL segment */
  slug: string;
}

export interface HyperExtractResult {
  tables: ParquetTable[];
}

/**
 * Extract all tables from a .hyper file into .parquet files.
 * Requires Python with tableauhyperapi and pyarrow installed.
 * Falls back gracefully if Python is unavailable.
 */
export async function extractToParquet(
  twbxOrHyperPath: string,
  outputDir: string,
  /** Optional slug prefix for output files — prevents name collisions across workbooks */
  slugPrefix?: string,
): Promise<HyperExtractResult> {
  await mkdir(outputDir, { recursive: true });

  const python = await findPython();
  if (!python) return scanExistingParquets(outputDir, slugPrefix);

  // .twbx is a ZIP — extract the .hyper file(s) to a temp dir first
  let hyperPath = twbxOrHyperPath;
  let tempDir: string | null = null;

  if (twbxOrHyperPath.toLowerCase().endsWith('.twbx')) {
    tempDir = await mkdtemp(path.join(tmpdir(), 'drexo-hyper-'));
    try {
      const zip = new AdmZip(twbxOrHyperPath);
      const hyperEntries = zip.getEntries().filter((e) => e.entryName.endsWith('.hyper'));
      if (hyperEntries.length === 0) {
        process.stderr.write('[drexo] No .hyper file found inside .twbx archive\n');
        return { tables: [] };
      }
      // Use the largest .hyper (main extract)
      const entry = hyperEntries.sort((a, b) => b.header.size - a.header.size)[0];
      hyperPath = path.join(tempDir, path.basename(entry.entryName));
      await writeFileAsync(hyperPath, entry.getData());
    } catch (e) {
      process.stderr.write(`[drexo] Failed to extract .hyper from .twbx: ${e}\n`);
      await rm(tempDir, { recursive: true, force: true });
      return { tables: [] };
    }
  }

  // Pass slug prefix to Python so it writes directly to the correct filename
  const result = await runPythonExtractor(python, hyperPath, outputDir, slugPrefix);

  if (tempDir) {
    await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }

  // If extraction produced nothing, reuse any pre-existing parquet files
  if (result.tables.length === 0) return scanExistingParquets(outputDir, slugPrefix);

  return result;
}

async function runPythonExtractor(
  python: string,
  hyperPath: string,
  outputDir: string,
  slugPrefix?: string,
): Promise<HyperExtractResult> {
  return new Promise((resolve) => {
    const args = [PYTHON_SCRIPT, hyperPath, outputDir];
    if (slugPrefix) args.push(slugPrefix);
    const child = spawn(python, args);

    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d) => { stdout += d; });
    child.stderr.on('data', (d) => { stderr += d; });

    child.on('close', (code) => {
      if (code !== 0) {
        const err = parseJson<{ error?: string }>(stderr) ?? { error: stderr.trim() };
        process.stderr.write(`[drexo] hyper extraction warning: ${err.error ?? 'unknown error'}\n`);
        resolve({ tables: [] });
        return;
      }

      const raw = parseJson<Record<string, string>>(stdout);
      if (!raw) {
        process.stderr.write('[drexo] hyper extraction: unexpected output from Python script\n');
        resolve({ tables: [] });
        return;
      }

      const tables: ParquetTable[] = Object.entries(raw).map(([qualifiedName, parquetPath]) => ({
        qualifiedName,
        parquetPath,
        slug: slugFromPath(parquetPath),
      }));

      resolve({ tables });
    });
  });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function findPython(): Promise<string | null> {
  for (const cmd of ['python3', 'python']) {
    const found = await testCommand(cmd, ['--version']);
    if (found) return cmd;
  }
  process.stderr.write(
    '[drexo] Python not found — skipping .hyper extraction.\n' +
    '        Install Python 3 with: pip install tableauhyperapi pyarrow\n'
  );
  return null;
}

function testCommand(cmd: string, args: string[]): Promise<boolean> {
  return new Promise((resolve) => {
    const child = spawn(cmd, args, { stdio: 'ignore' });
    child.on('error', () => resolve(false));
    child.on('close', (code) => resolve(code === 0));
  });
}

function parseJson<T>(text: string): T | null {
  try { return JSON.parse(text.trim()) as T; }
  catch { return null; }
}

function slugFromPath(filePath: string): string {
  return path.basename(filePath, '.parquet');
}

/** Return any .parquet files already present in outputDir (optionally filtered by slugPrefix). */
async function scanExistingParquets(outputDir: string, slugPrefix?: string): Promise<HyperExtractResult> {
  if (!existsSync(outputDir)) return { tables: [] };
  let entries: string[];
  try { entries = await readdir(outputDir); } catch { return { tables: [] }; }
  const tables: ParquetTable[] = entries
    .filter((f) => f.endsWith('.parquet') && (!slugPrefix || f.startsWith(slugPrefix)))
    .map((f) => {
      const absPath = path.join(outputDir, f);
      const slug = slugFromPath(absPath);
      return { qualifiedName: `extract.${slug}`, parquetPath: absPath, slug };
    });
  if (tables.length > 0) {
    process.stderr.write(`[drexo] Python unavailable — reusing ${tables.length} existing parquet file(s) from ${outputDir}\n`);
  }
  return { tables };
}
