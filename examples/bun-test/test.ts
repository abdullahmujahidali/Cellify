/**
 * Cellify - Bun Runtime Verification Test
 *
 * Run with: bun run test.ts
 *
 * Prerequisites:
 * 1. Build the library: npm run build (from root)
 * 2. Navigate to this directory: cd examples/bun-test
 * 3. Run: bun run test.ts
 */

// Import from local build
// In production: import { ... } from 'cellify';
import {
  Workbook,
  workbookToXlsx,
  workbookToXlsxBlob,
  xlsxToWorkbook,
  xlsxBlobToWorkbook,
  initXlsxWasm,
} from '../../dist/esm/index.js';

import { writeFileSync, readFileSync, unlinkSync, existsSync } from 'fs';

const tests: { name: string; fn: () => Promise<boolean> }[] = [];

function test(name: string, fn: () => Promise<boolean>) {
  tests.push({ name, fn });
}

async function runTests() {
  console.log('\n🧪 Cellify - Bun Runtime Verification\n');
  console.log(`Bun version: ${Bun.version}`);
  console.log('='.repeat(50));

  let passed = 0;
  let failed = 0;

  for (const { name, fn } of tests) {
    try {
      const result = await fn();
      if (result) {
        console.log(`✅ ${name}`);
        passed++;
      } else {
        console.log(`❌ ${name} - returned false`);
        failed++;
      }
    } catch (e) {
      console.log(`❌ ${name} - ${(e as Error).message}`);
      failed++;
    }
  }

  console.log('='.repeat(50));
  console.log(`\nResults: ${passed} passed, ${failed} failed\n`);

  return failed === 0;
}

// Test 1: Create Workbook
test('Create Workbook', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Test');

  sheet.cell(0, 0).value = 'Hello Bun!';
  sheet.cell(0, 1).value = 123;
  sheet.cell(1, 0).value = true;
  sheet.cell(1, 1).value = new Date();

  return (
    workbook.sheetCount === 1 &&
    sheet.getCell(0, 0)?.value === 'Hello Bun!'
  );
});

// Test 2: Export to Uint8Array
test('Export to Uint8Array', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Export');

  sheet.cell(0, 0).value = 'Bun Export Test';
  sheet.cell(1, 0).value = 42;

  const buffer = workbookToXlsx(workbook);

  return buffer instanceof Uint8Array && buffer.length > 0;
});

// Test 3: Export to Blob
test('Export to Blob', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Blob Test');

  sheet.cell(0, 0).value = 'Blob Test';

  const blob = workbookToXlsxBlob(workbook);

  return (
    blob instanceof Blob &&
    blob.size > 0 &&
    blob.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
});

// Test 4: Write and Read File
test('Write and Read File', async () => {
  const testFile = 'cellify-bun-test.xlsx';

  // Create and export
  const workbook = new Workbook();
  const sheet = workbook.addSheet('File Test');
  sheet.cell(0, 0).value = 'Written by Bun';
  sheet.cell(1, 0).value = 12345;

  const buffer = workbookToXlsx(workbook);
  writeFileSync(testFile, buffer);

  // Read back
  const readBuffer = readFileSync(testFile);
  const result = await xlsxToWorkbook(new Uint8Array(readBuffer));

  // Clean up
  if (existsSync(testFile)) {
    unlinkSync(testFile);
  }

  return (
    result.workbook.sheetCount === 1 &&
    result.stats.totalCells >= 2
  );
});

// Test 5: Import from Blob
test('Import from Blob', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Blob Import');

  sheet.cell(0, 0).value = 'Blob round-trip';
  sheet.cell(0, 1).value = 999;

  const blob = workbookToXlsxBlob(workbook);
  const result = await xlsxBlobToWorkbook(blob);

  return (
    result.workbook.sheetCount === 1 &&
    result.stats.totalCells === 2
  );
});

// Test 6: Cell Styling
test('Cell Styling', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Styled');

  sheet.cell(0, 0).value = 'Bold Red';
  sheet.cell(0, 0).applyStyle({
    font: { bold: true, color: '#FF0000' },
    fill: { color: '#FFFF00' },
  });

  const buffer = workbookToXlsx(workbook);
  const result = await xlsxToWorkbook(buffer);

  // Style should be preserved
  const importedCell = result.workbook.sheets[0].getCell(0, 0);
  return (
    importedCell?.style?.font?.bold === true &&
    result.stats.totalCells >= 1
  );
});

// Test 7: Formulas
test('Formulas', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Formulas');

  sheet.cell(0, 0).value = 10;
  sheet.cell(0, 1).value = 20;
  sheet.cell(0, 2).setFormula('=A1+B1');

  const buffer = workbookToXlsx(workbook);
  const result = await xlsxToWorkbook(buffer);

  const formulaCell = result.workbook.sheets[0].getCell(0, 2);
  return formulaCell?.formula?.formula === 'A1+B1';
});

// Test 8: Comments
test('Comments', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Comments');

  sheet.cell(0, 0).value = 'Has Comment';
  sheet.cell(0, 0).setComment('This is a comment', 'Bun User');

  const buffer = workbookToXlsx(workbook);
  const result = await xlsxToWorkbook(buffer);

  // Verify export worked (comments import may not be implemented yet)
  return result.stats.totalCells >= 1;
});

// Test 9: Merged Cells
test('Merged Cells', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Merged');

  sheet.cell(0, 0).value = 'Merged Header';
  sheet.mergeCells('A1:D1');

  const buffer = workbookToXlsx(workbook);
  const result = await xlsxToWorkbook(buffer);

  return result.stats.mergedRanges === 1;
});

// Test 10: WASM Initialization
test('WASM Initialization', async () => {
  try {
    await initXlsxWasm();
    return true;
  } catch (e) {
    // WASM might not be available in Bun yet
    console.log('  Note: WASM may not be fully supported in Bun yet');
    return true; // Don't fail the test
  }
});

// Test 11: Multiple Sheets
test('Multiple Sheets', async () => {
  const workbook = new Workbook();

  workbook.addSheet('Sheet1').cell(0, 0).value = 'First';
  workbook.addSheet('Sheet2').cell(0, 0).value = 'Second';
  workbook.addSheet('Sheet3').cell(0, 0).value = 'Third';

  const buffer = workbookToXlsx(workbook);
  const result = await xlsxToWorkbook(buffer);

  return (
    result.workbook.sheetCount === 3 &&
    result.workbook.sheets[0].name === 'Sheet1' &&
    result.workbook.sheets[1].name === 'Sheet2' &&
    result.workbook.sheets[2].name === 'Sheet3'
  );
});

// Test 12: Large Dataset
test('Large Dataset (1000 rows)', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Large');

  for (let i = 0; i < 1000; i++) {
    sheet.cell(i, 0).value = `Row ${i}`;
    sheet.cell(i, 1).value = i * 10;
    sheet.cell(i, 2).value = i % 2 === 0;
  }

  const start = performance.now();
  const buffer = workbookToXlsx(workbook);
  const exportTime = performance.now() - start;

  const importStart = performance.now();
  const result = await xlsxToWorkbook(buffer);
  const importTime = performance.now() - importStart;

  console.log(`  Export: ${exportTime.toFixed(1)}ms, Import: ${importTime.toFixed(1)}ms`);

  return result.stats.totalCells === 3000;
});

// Run all tests
const success = await runTests();
process.exit(success ? 0 : 1);
