import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { generateEconomicTree, generateFunctionalTree, hasDecimals } from './prepare-data-lib.ts';
import { validateTotals } from './validate-data.js';
import type { BudgetNode } from '../src/utils/types.ts';

type Row = [string | number, string, string | number, ...(string | number)[]];

function sheetOf(rows: Row[], header: string[] = ['#', 'Megnevezés', 'Összesen']) {
	const wb = new ExcelJS.Workbook();
	const ws = wb.addWorksheet('2026 KIADÁS');
	ws.getCell('A1').value = '2026 KIADÁS';
	ws.getRow(2).values = header;
	rows.forEach((r, i) => (ws.getRow(3 + i).values = r));
	return ws;
}

function captureWarnings(fn: () => void) {
	const messages: string[] = [];
	const original = console.error;
	console.error = (...args: unknown[]) => messages.push(args.join(' '));
	try {
		fn();
	} finally {
		console.error = original;
	}
	return messages;
}

function find(node: BudgetNode, id: string): BudgetNode | undefined {
	if (node.id === id) return node;
	for (const c of node.children || []) {
		const f = find(c, id);
		if (f) return f;
	}
}

const FUNC_TREE = () =>
	({
		1: { id: 1, name: 'Általános közszolgáltatások', parent: null },
		11130: { id: 11130, name: 'Önkormányzatok jogalkotó tev.', parent: 1 },
		10: { id: 10, name: 'Szociális védelem', parent: null },
		1070: { id: 1070, name: 'Egyéb szociális kirekesztettség', parent: 10 },
		107051: { id: 107051, name: 'Szociális étkeztetés', parent: 1070 },
		107052: { id: 107052, name: 'Házi segítségnyújtás', parent: 1070 },
	}) as unknown as Record<number, BudgetNode>;

test('hasDecimals: decimals vs thousand separators', () => {
	assert.equal(hasDecimals(1234.5), true);
	assert.equal(hasDecimals(1234), false);
	assert.equal(hasDecimals('1 234,5'), true);
	assert.equal(hasDecimals('1.25'), true);
	assert.equal(hasDecimals('1,234,567'), false);
	assert.equal(hasDecimals('1 234 567'), false);
});

test('a well-formed sheet produces no warnings', () => {
	const ws = sheetOf([
		[10, 'Személyi juttatások (K1)', 600],
		[11, 'Foglalkoztatottak (K11)', 500],
		[12, 'Külső személyi juttatások (K12)', 100],
	]);
	let root!: BudgetNode;
	const warnings = captureWarnings(() => (root = generateEconomicTree(ws)));
	assert.deepEqual(warnings, []);
	assert.equal(root.value, 600);
});

test('warns about a skipped row with a single-digit first column', () => {
	const ws = sheetOf([
		[10, 'Személyi juttatások (K1)', 600],
		[11, 'Foglalkoztatottak (K11)', 500],
		[5, 'Külső személyi juttatások (K12)', 100],
	]);
	let root!: BudgetNode;
	const warnings = captureWarnings(() => (root = generateEconomicTree(ws)));
	assert.equal(find(root, 'K12'), undefined); // behaviour unchanged: the row is skipped
	assert.equal(warnings.length, 1);
	assert.match(warnings[0]!, /5\. sor kimarad/);
});

test('warns about a decimal amount', () => {
	const ws = sheetOf([
		[10, 'Dologi kiadások (K3)', 400.5],
		[11, 'Készletbeszerzés (K31)', '400,5'],
	]);
	let root!: BudgetNode;
	const warnings = captureWarnings(() => (root = generateEconomicTree(ws)));
	assert.equal(find(root, 'K3')!.value, 4005); // behaviour unchanged
	assert.equal(warnings.length, 2);
	assert.match(warnings[0]!, /nem egész szám/);
});

test('warns when a node is attached to a non-prefix parent', () => {
	const ws = sheetOf([
		[10, 'Dologi kiadások (K3)', 400],
		[11, 'Készletbeszerzés (K31)', 400],
		[12, 'Tartalékok (K512)', 50],
	]);
	let root!: BudgetNode;
	const warnings = captureWarnings(() => (root = generateEconomicTree(ws)));
	assert.ok(find(find(root, 'K31')!, 'K512')); // behaviour unchanged
	assert.equal(warnings.length, 1);
	assert.match(warnings[0]!, /K512 szülője nem létezik, ide került: K31/);
});

test('warns about a negative functional value and a non-leaf function code', () => {
	const ws = sheetOf(
		[
			[10, 'Személyi juttatások (K1)', 600],
			[11, 'Dologi kiadások (K3)', 400],
			[99, 'Kiadások összesen', 1000, 600, 300, 700, -100],
		],
		[
			'#',
			'Megnevezés',
			'Összesen',
			'011130 Jogalkotó',
			'107051 Étkeztetés',
			'1070 Nem levél',
			'107052 Házi',
		],
	);
	let func!: BudgetNode | null;
	const warnings = captureWarnings(() => (func = generateFunctionalTree(ws, FUNC_TREE())));
	assert.ok(func);
	assert.equal(find(func!, 1070 as unknown as string)?.value, 400); // 300 + |-100|, the 700 is lost
	assert.equal(warnings.length, 2);
	assert.match(warnings.join('\n'), /107052 funkcionális érték negatív/);
	assert.match(warnings.join('\n'), /1070 funkcionális kategóriának alkategóriái vannak/);
});

test('validateTotals warns when functional and economic totals differ', () => {
	const tree = (value: number) => ({ name: 'Összesen', value, children: [] });
	const warnings = captureWarnings(() =>
		validateTotals({
			'2025': {
				expense: { econ: tree(1000), func: tree(1000) },
				income: { econ: tree(1000), func: null },
			},
			'2026': {
				expense: { econ: tree(1000), func: tree(900) },
				income: { econ: tree(1000), func: null },
			},
		}),
	);
	assert.equal(warnings.filter((w) => /végösszeg eltér/.test(w)).length, 1);
	assert.match(warnings.join('\n'), /2026, expense/);
});
