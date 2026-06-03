import { existsSync } from 'node:fs';
import { writeFile, mkdir, readdir, stat } from 'node:fs/promises';
import path from 'node:path';

import chalk from 'chalk';
import { Command } from 'commander';

import { parseWorkbook } from './parsers/index.js';
import { generateMarkdownModel } from './generators/index.js';
import { generateVisualBlueprint } from './generators/visual-blueprint.js';
import { generateLayoutHtml } from './generators/layout-html.js';
import { scaffoldStaticFiles } from './generators/react/scaffolder.js';
import { generateDashboardComponent, type DashboardGeneratorResult } from './generators/react/dashboard-component.js';
import { generateApiServer } from './generators/react/api-server.js';
import { extractToParquet } from './extractors/hyper-extractor.js';
import { generateAppTsx, generateWorkbookListTsx } from './generators/react/app-shell.js';
import { generateReportIndex, toSubdirName, type IndexEntry } from './generators/report-index.js';
import { enrichWorkbook, type EnrichStep } from './enrichers/index.js';
import { logger as log } from './utils/logger.js';

declare const __DREXO_VERSION__: string;

const program = new Command();

program
  .name('drexo')
  .description(
    chalk.cyan(
      'Migrate Tableau workbooks (.twb / .twbx) into modern React dashboards.\n' +
        'v0.1 ships the metadata layer; React generation lands in v0.2.'
    )
  )
  .version(__DREXO_VERSION__, '-v, --version', 'output the current version')
  .helpOption('-h, --help', 'display help for command')
  .option('-d, --debug', 'enable debug logging');

program.on('option:debug', () => {
  process.env.DEBUG = '1';
});

program.addHelpText(
  'after',
  `
${chalk.bold('Examples:')}
  $ drexo analyze report.twbx
  $ drexo analyze report.twbx --output-dir ./reports
  $ drexo analyze ./examples/                        # all .twbx in directory
  $ drexo analyze a.twbx b.twbx --output-dir reports
  $ drexo analyze report.twbx --enrich               # AI-enriched output

${chalk.bold('Learn more:')}
  https://github.com/raguvindtharanitharan/drexo
`
);

// ---------------------------------------------------------------------------
// analyze
// ---------------------------------------------------------------------------

program
  .command('analyze <files...>')
  .description(
    'Parse one or more Tableau workbooks (.twbx) and write metadata files.\n' +
    'Accepts files, glob patterns, or a directory path.'
  )
  .option('--output-dir <path>', 'write all outputs into an organized directory (one subdir per workbook)')
  .option('--enrich', 'enrich metadata with AI-generated summaries and descriptions (requires ANTHROPIC_API_KEY)')
  .option('--enrich-model <model>', 'Claude model to use for enrichment', 'claude-haiku-4-5-20251001')
  .action(async (
    rawArgs: string[],
    options: { outputDir?: string; enrich?: boolean; enrichModel?: string },
  ) => {
    if (options.enrich && !process.env.ANTHROPIC_API_KEY) {
      console.error(chalk.red('\n✖ --enrich requires ANTHROPIC_API_KEY\n'));
      console.error(chalk.gray('  Set the environment variable and re-run:'));
      console.error(chalk.cyan('    export ANTHROPIC_API_KEY=sk-ant-...'));
      console.error(chalk.cyan(`    drexo analyze ${rawArgs[0]} --enrich\n`));
      console.error(chalk.gray('  Get an API key at: https://console.anthropic.com\n'));
      process.exit(1);
    }

    // Collect all .twbx files from args (each may be a file or directory)
    const files = await collectTwbxFiles(rawArgs);
    if (files.length === 0) {
      log.error('No .twbx files found.');
      process.exit(1);
    }

    const isBatch = files.length > 1 || options.outputDir !== undefined;
    if (isBatch) {
      log.info(`📂 Found ${files.length} workbook${files.length === 1 ? '' : 's'}`);
    }

    // Prepare output directory if needed
    if (options.outputDir) {
      await mkdir(options.outputDir, { recursive: true });
    }

    // Process with concurrency cap of 3
    const indexEntries: IndexEntry[] = [];
    const totalStart = performance.now();

    await runConcurrently(files, 3, async (file) => {
      const entry = await analyzeOne(file, options, isBatch);
      if (entry) indexEntries.push(entry);
    });

    // Write index.md if using output-dir with multiple workbooks
    if (options.outputDir && indexEntries.length > 1) {
      const indexPath = path.join(options.outputDir, 'index.md');
      const indexContent = generateReportIndex(indexEntries, new Date().toISOString().split('T')[0]);
      await writeFile(indexPath, indexContent, 'utf-8');
      log.success(`Wrote ${chalk.cyan(indexPath)} (catalog of ${indexEntries.length} workbooks)`);
    }

    if (isBatch) {
      const elapsed = ((performance.now() - totalStart) / 1000).toFixed(1);
      log.success(`Done — ${files.length} workbook${files.length === 1 ? '' : 's'} in ${elapsed}s`);
    }
  });

// ---------------------------------------------------------------------------
// Core: analyze one workbook
// ---------------------------------------------------------------------------

async function analyzeOne(
  file: string,
  options: { outputDir?: string; enrich?: boolean; enrichModel?: string },
  quiet = false,
): Promise<IndexEntry | null> {
  const start = performance.now();
  const basename = path.basename(file);
  const prefix = quiet ? chalk.gray(`[${basename}] `) : '';

  try {
    if (!existsSync(file)) {
      log.error(`${prefix}File not found: ${file}`);
      return null;
    }

    log.info(`🔍 ${prefix}Analyzing ${chalk.bold(basename)}`);

    let workbook = await parseWorkbook(file);
    log.success(
      `${prefix}Parsed ${workbook.dataSources.length} data source${workbook.dataSources.length === 1 ? '' : 's'}, ` +
      `${workbook.worksheets.length} worksheets, ` +
      `${workbook.dashboards.length} dashboards`
    );

    let visualBlueprint = generateVisualBlueprint(workbook);
    let enriched = false;

    if (options.enrich) {
      log.info(`${prefix}✦ Enriching (${workbook.worksheets.length} worksheets, ${workbook.calculations.length} calculations)...`);

      const { workbook: enrichedWb, enrichedVisualBlueprint, usage } = await enrichWorkbook(workbook, {
        apiKey: process.env.ANTHROPIC_API_KEY!,
        model: options.enrichModel,
        visualBlueprint,
        onProgress: (step: EnrichStep) => {
          const label: Record<EnrichStep['phase'], string> = {
            executive_summary: 'executive summary',
            worksheet_descriptions: 'worksheet descriptions',
            calc_simplifications: 'calc simplifications',
            visual_blueprint: 'visual blueprint',
          };
          const cacheNote = step.fromCache > 0 ? chalk.gray(` (${step.fromCache} from cache)`) : '';
          if (step.done === step.total) {
            const counter = step.total > 1 ? ` [${step.done}/${step.total}]` : '';
            log.success(`${prefix}  ${label[step.phase]}${counter}${cacheNote}`);
          }
        },
      });

      workbook = enrichedWb;
      if (enrichedVisualBlueprint) visualBlueprint = enrichedVisualBlueprint;
      enriched = true;

      const costDisplay = usage.estimatedCostUsd < 0.001
        ? chalk.gray('< $0.001')
        : chalk.gray(`~$${usage.estimatedCostUsd.toFixed(4)}`);
      log.info(`${prefix}  Tokens: ${usage.inputTokens.toLocaleString()} in / ${usage.outputTokens.toLocaleString()} out  ${costDisplay}`);
    }

    // Determine output paths
    const { modelPath, visualPath } = resolveOutputPaths(file, options.outputDir);
    await mkdir(path.dirname(modelPath), { recursive: true });

    const markdown = generateMarkdownModel(workbook);
    await writeFile(modelPath, markdown, 'utf-8');
    await writeFile(visualPath, visualBlueprint, 'utf-8');

    const elapsed = ((performance.now() - start) / 1000).toFixed(2);
    log.success(`${prefix}Wrote ${chalk.cyan(modelPath)} (${(markdown.length / 1024).toFixed(1)} KB) in ${elapsed}s`);
    log.success(`${prefix}Wrote ${chalk.cyan(visualPath)} (${(visualBlueprint.length / 1024).toFixed(1)} KB)`);

    if (!options.outputDir) return null; // no index needed for single-file mode

    return {
      workbook,
      subdirName: toSubdirName(path.basename(file)),
      enriched,
    };
  } catch (e) {
    const reason = e instanceof Error ? e.message : String(e);
    log.error(`${prefix}Could not analyze: ${reason}`);
    if (process.env.DEBUG && e instanceof Error && e.stack) {
      console.error(chalk.gray(e.stack));
    }
    return null;
  }
}

// ---------------------------------------------------------------------------
// Output path resolution
// ---------------------------------------------------------------------------

function resolveOutputPaths(
  inputFile: string,
  outputDir: string | undefined,
): { modelPath: string; visualPath: string } {
  if (outputDir) {
    const subdir = toSubdirName(path.basename(inputFile));
    const base = path.join(outputDir, subdir);
    return {
      modelPath: path.join(base, 'model.md'),
      visualPath: path.join(base, 'visual.md'),
    };
  }
  // Legacy single-file behavior: alongside input
  const parsed = path.parse(inputFile);
  const base = path.join(parsed.dir, parsed.name);
  return {
    modelPath: `${base}.model.md`,
    visualPath: `${base}.visual.md`,
  };
}

// ---------------------------------------------------------------------------
// File collection — expand directories and filter for .twbx
// ---------------------------------------------------------------------------

async function collectTwbxFiles(args: string[]): Promise<string[]> {
  const files: string[] = [];
  for (const arg of args) {
    if (!existsSync(arg)) {
      log.warn(`Skipping — not found: ${arg}`);
      continue;
    }
    const s = await stat(arg);
    if (s.isDirectory()) {
      const found = await findTwbx(arg);
      files.push(...found);
    } else if (arg.endsWith('.twbx') || arg.endsWith('.twb')) {
      files.push(arg);
    } else {
      log.warn(`Skipping — not a .twbx file: ${arg}`);
    }
  }
  return [...new Set(files)]; // deduplicate
}

async function findTwbx(dir: string): Promise<string[]> {
  const results: string[] = [];
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...await findTwbx(full));
    } else if (entry.name.endsWith('.twbx') || entry.name.endsWith('.twb')) {
      results.push(full);
    }
  }
  return results;
}

// ---------------------------------------------------------------------------
// Concurrency helper — runs tasks with a max parallelism cap
// ---------------------------------------------------------------------------

async function runConcurrently<T>(
  items: T[],
  limit: number,
  fn: (item: T) => Promise<void>,
): Promise<void> {
  const queue = [...items];
  const workers = Array.from(
    { length: Math.min(limit, items.length) },
    async () => {
      while (queue.length > 0) {
        const item = queue.shift()!;
        await fn(item);
      }
    }
  );
  await Promise.all(workers);
}

// ---------------------------------------------------------------------------
// migrate
// ---------------------------------------------------------------------------

program
  .command('migrate <files...>')
  .description('Generate a single React app from one or more Tableau workbooks. Accepts files or a directory.')
  .option('--output-dir <path>', 'output directory (default: ./drexo-app)')
  .option('--enrich', 'enrich metadata with AI before migrating (requires ANTHROPIC_API_KEY)')
  .action(async (rawArgs: string[], options: { outputDir?: string; enrich?: boolean }) => {
    const start = performance.now();
    try {
      const files = await collectTwbxFiles(rawArgs);
      if (files.length === 0) { log.error('No .twbx files found.'); process.exit(1); }

      const outputDir = options.outputDir ?? path.join(process.cwd(), 'drexo-app');
      await mkdir(outputDir, { recursive: true });
      await mkdir(path.join(outputDir, 'src', 'workbooks'), { recursive: true });

      log.info(`⚛  Migrating ${files.length} workbook${files.length === 1 ? '' : 's'} → ${chalk.cyan(outputDir)}`);

      // Parse all workbooks (auto-analyze if md files missing)
      const workbooks = [];
      for (const file of files) {
        const parsed = path.parse(file);
        const modelPath = path.join(parsed.dir, `${parsed.name}.model.md`);
        const visualPath = path.join(parsed.dir, `${parsed.name}.visual.md`);
        if (!existsSync(modelPath) || !existsSync(visualPath)) {
          log.info(`  ↳ Analyzing ${chalk.bold(parsed.base)} first...`);
          await analyzeOne(file, { enrich: options.enrich ?? false });
        }
        const workbook = await parseWorkbook(file);
        const slug = toSubdirName(parsed.base);

        // Extract .hyper → .parquet (requires Python + tableauhyperapi + pyarrow)
        const parquetDir = path.join(outputDir, 'api', 'data');
        log.info(`  ↳ Extracting data from ${chalk.bold(parsed.base)}...`);
        const extraction = await extractToParquet(file, parquetDir, slug);
        if (extraction.tables.length > 0) {
          log.success(`  Extracted ${extraction.tables.length} table(s) → api/data/`);
        } else {
          log.warn(`  No data extracted — API will return empty results. Install: pip install tableauhyperapi pyarrow`);
        }

        workbooks.push({ file, workbook, slug, parquetTables: extraction.tables });
      }

      // Scaffold static files
      const appTitle = workbooks.length === 1
        ? workbooks[0].workbook.metadata.name
        : 'drexo Dashboards';
      for (const sf of scaffoldStaticFiles(appTitle)) {
        const dest = path.join(outputDir, sf.relativePath);
        await mkdir(path.dirname(dest), { recursive: true });
        await writeFile(dest, sf.content, 'utf-8');
      }

      // Generate per-workbook Dashboard.tsx + layout-preview.html
      const entries = [];
      for (const { workbook, slug, parquetTables } of workbooks) {
        const dashDir = path.join(outputDir, 'src', 'workbooks', slug);
        await mkdir(dashDir, { recursive: true });
        const { serverTs: _s, specs: querySpecs } = generateApiServer(workbook, parquetTables, slug);
        const { dashboardTsx, dataFiles } = generateDashboardComponent(workbook, slug, querySpecs);
        await writeFile(path.join(dashDir, 'Dashboard.tsx'), dashboardTsx, 'utf-8');
        for (const df of dataFiles) {
          const dest = path.join(outputDir, df.relativePath);
          await mkdir(path.dirname(dest), { recursive: true });
          await writeFile(dest, df.content, 'utf-8');
        }
        const previewHtml = generateLayoutHtml(workbook);
        await writeFile(path.join(dashDir, 'layout-preview.html'), previewHtml, 'utf-8');
        entries.push({ workbook, slug });
        log.success(`  ${chalk.cyan(`src/workbooks/${slug}/Dashboard.tsx`)} (${dataFiles.length} data file${dataFiles.length === 1 ? '' : 's'})`);
        log.success(`  ${chalk.cyan(`src/workbooks/${slug}/layout-preview.html`)}`);
      }

      // Generate App.tsx + WorkbookList.tsx
      await writeFile(path.join(outputDir, 'src', 'App.tsx'), generateAppTsx(entries), 'utf-8');
      await writeFile(path.join(outputDir, 'src', 'WorkbookList.tsx'), generateWorkbookListTsx(entries), 'utf-8');

      // Generate one combined API server covering all workbooks
      await mkdir(path.join(outputDir, 'api'), { recursive: true });
      const allParquetTables = workbooks.flatMap((w) => w.parquetTables);
      // Use first workbook to generate the server (sources list covers all)
      if (workbooks.length > 0) {
        const { serverTs } = generateApiServer(workbooks[0].workbook, allParquetTables, workbooks[0].slug);
        await writeFile(path.join(outputDir, 'api', 'server.ts'), serverTs, 'utf-8');
        log.success(`  ${chalk.cyan('api/server.ts')} (Express + DuckDB, ${allParquetTables.length} source${allParquetTables.length === 1 ? '' : 's'})`);
      }

      const elapsed = ((performance.now() - start) / 1000).toFixed(2);
      log.success(`Done in ${elapsed}s`);
      log.info(`\n  ${chalk.bold('Next steps:')}`);
      log.info(`    cd ${outputDir}`);
      log.info(`    npm install`);
      log.info(`    npm run dev`);
    } catch (e) {
      const reason = e instanceof Error ? e.message : String(e);
      log.error(`Migration failed: ${reason}`);
      if (process.env.DEBUG && e instanceof Error && e.stack) console.error(chalk.gray((e as Error).stack));
      process.exit(1);
    }
  });

// ---------------------------------------------------------------------------
// No-command → help
// ---------------------------------------------------------------------------

if (process.argv.length === 2) {
  program.help();
}

program.parse(process.argv);
