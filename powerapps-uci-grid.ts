// libs/powerapps-uci-grid.ts
import { Page, Locator, expect } from '@playwright/test';

/**
 * Options describing how to locate the target row.
 * Either provide `logicalName` (preferred) OR `columnHeader` (fallback).
 */
export type GridCellLocator = {
  /** Cell text to match (case-insensitive, exact by default). */
  value: string | RegExp;

  /** Dynamics logical name of the column, e.g., "name", "accountnumber", "emailaddress1". */
  logicalName?: string;

  /** Visible column header text, e.g., "Name", "Account Number". Used if logicalName is unknown. */
  columnHeader?: string;

  /** If provided, do a partial (substring) match instead of exact; ignored when `value` is a RegExp. */
  partial?: boolean;
};

type FindMode = 'open-record' | 'select-row';

const rx = (v: string | RegExp, partial?: boolean) =>
  v instanceof RegExp ? v : new RegExp(partial ? escapeForRx(v) : `^${escapeForRx(v)}$`, 'i');

function escapeForRx(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function scrollIntoViewIfNeeded(loc: Locator) {
  // Try native scroll into view; if virtualization hides it, this will still help when it renders.
  await loc.evaluateAll((els) => els.forEach((el) => el.scrollIntoView({ block: 'center' }))).catch(() => {});
}

/**
 * Resolve the column's gridcell locator either by logical name or by header title.
 */
function gridCellLocator(page: Page, grid: Locator, q: GridCellLocator): Locator {
  if (q.logicalName) {
    // Fast path: Dynamics gives data-id like: cell-<logicalName>
    return grid.locator(`[role="gridcell"][data-id="cell-${q.logicalName}"]`);
  }

  // Fallback: map header -> column index, then pick the Nth cell in each row
  const header = grid.getByRole('columnheader', { name: rx(q.columnHeader ?? '', false) }).first();
  const rows = grid.locator('[role="row"]').filter({ has: grid.getByRole('gridcell') });

  const nthCellInEachRow = rows.locator('> [role="gridcell"]').locator((_, i) => {
    // This dummy filter is replaced later; we compute index in evaluate below.
    return i === 0;
  });

  // We can’t compute index in pure locator expressions portably, so we return a locator that we’ll refine
  // by filtering with `has=` afterwards. Practically, we’ll just use by header fast path if possible.
  // To keep this function synchronous, we approximate by targeting all gridcells and then refining by :has(header match)
  // However, to keep it robust, we’ll instead select all gridcells and filter by header via aria-colindex mapping.

  // Safer generic locator: any gridcell whose aria-colindex equals that of the header
  const colIndex = header.getAttribute('aria-colindex');
  return grid.locator(`[role="gridcell"][aria-colindex="${colIndex}"]`);
}

/**
 * Find the row that contains a cell matching the requested column + value.
 * This auto-scrolls through a virtualized grid until the target is found or maxScrolls reached.
 */
async function findRowByCell(page: Page, q: GridCellLocator, { maxScrolls = 30 } = {}) {
  const grid = page.getByRole('grid').first(); // main data grid on the page
  await expect(grid, 'Grid should be visible').toBeVisible();

  const cellColumn = gridCellLocator(page, grid, q);
  const valueRx = rx(q.value, q.partial);

  let attempts = 0;
  while (attempts <= maxScrolls) {
    // Look for any cell in that column whose text matches value
    const candidateCell = cellColumn.filter({ hasText: valueRx }).first();

    if (await candidateCell.isVisible().catch(() => false)) {
      // Found a visible cell; return its containing row
      const row = candidateCell.locator('xpath=ancestor::*[@role="row"][1]');
      await expect(row, 'Row for matching cell should be visible').toBeVisible();
      return row;
    }

    // Not visible yet – try to scroll the grid body a bit to load more rows
    const gridBody = grid.locator('xpath=.//*[contains(@class,"wj-cells") or contains(@class,"viewport") or @role="rowgroup"]').first();
    if (await gridBody.count() === 0) {
      // Fallback: scroll the page
      await page.mouse.wheel(0, 1200);
    } else {
      await gridBody.evaluate((el) => { (el as HTMLElement).scrollBy(0, 800); }).catch(() => {});
    }

    // Also try asking Playwright to look again after scroll
    await scrollIntoViewIfNeeded(cellColumn);

    attempts += 1;
  }

  throw new Error(
    `Could not find a row with ${q.logicalName ? `logicalName="${q.logicalName}"` : `header="${q.columnHeader}"`} matching value "${q.value}" after ${maxScrolls} scrolls.`,
  );
}

/**
 * Opens the record by clicking the main clickable cell in the row (often the "primary name" column).
 */
export async function openRecordByCell(page: Page, q: GridCellLocator) {
  const row = await findRowByCell(page, q);
  // UCI usually renders the primary column as a link/button inside the row; if not, just click the row.
  const linkish = row.locator('a,button,[role="link"]').first();
  if (await linkish.isVisible().catch(() => false)) {
    await linkish.click();
  } else {
    await row.click({ position: { x: 40, y: 10 } }).catch(async () => await row.click());
  }
  // Assert a form opened (optional heuristic)
  await expect(page.getByRole('main')).toBeVisible();
}

/**
 * Selects the row via the leftmost selection control (checkbox/toggle button), without opening it.
 */
export async function selectRowByCell(page: Page, q: GridCellLocator) {
  const row = await findRowByCell(page, q);

  // Common patterns for the row-selector:
  // 1) <input type="checkbox">, 2) role="checkbox", 3) a button with title "Select row"
  const checkbox = row.getByRole('checkbox').first();
  if (await checkbox.isVisible().catch(() => false)) {
    // Some D365 checkboxes need .click() instead of .check() due to custom widgets
    await checkbox.click({ force: true });
    return;
  }

  const explicitButton = row.getByRole('button', { name: /select row/i }).first();
  if (await explicitButton.isVisible().catch(() => false)) {
    await explicitButton.click();
    return;
  }

  // Fallback: click the selector cell (usually the first gridcell)
  const firstCell = row.locator('> [role="gridcell"]').first();
  await firstCell.click();
}

/**
 * Bonus: quick "Search this view" filter before targeting, for faster finds on large views.
 * Call this to narrow the grid, then call openRecordByCell or selectRowByCell.
 */
export async function searchThisView(page: Page, text: string) {
  const search = page.getByPlaceholder(/search this view/i).first();
  if (await search.isVisible().catch(() => false)) {
    await search.fill(text);
    await search.press('Enter');
    // Wait for grid to refresh
    const grid = page.getByRole('grid').first();
    await expect(grid).toBeVisible();
  }
}
