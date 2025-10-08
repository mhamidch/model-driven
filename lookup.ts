// libs/uci-lookups.ts
import { Page, Locator, expect } from '@playwright/test';

/** Escape a string for use inside a RegExp */
const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Build a case-insensitive exact-match regex for an accessible name */
const labelRx = (label: string) => new RegExp(`^${esc(label)}$`, 'i');

/** Find a lookup input by its field label (Model-Driven Apps expose as textbox/combobox) */
async function findLookup(page: Page, label: string) {
  const cb = page.getByRole(/^(combobox|textbox)$/i, { name: labelRx(label) }).first();
  await expect(cb, `Lookup "${label}" should exist`).toBeVisible();
  return cb;
}

/** Try to remove any existing chips/pills/value from a lookup input */
async function clearLookupValue(input: Locator) {
  try { await input.press('Control+A'); } catch {}
  try { await input.press('Backspace'); } catch {}
  // Some controls ignore Backspace until focused via click
  try { await input.click({ position: { x: 5, y: 10 } }); } catch {}
  try { await input.fill(''); } catch {}
}

/** Robust wait for Dataverse search calls to return (reduces flakiness) */
async function waitForDataverseSearch(page: Page, timeout: number) {
  // Typical lookup search calls hit /api/data/v9.x/... endpoints
  // We wait for at least one response after typing/opening the list.
  await page.waitForResponse(
    (res) => {
      const url = res.url();
      const ok = /\/api\/data\/v9\.\d+\/.+/i.test(url);
      // Only count successful responses to avoid returning too early
      return ok && res.status() < 500;
    },
    { timeout }
  ).catch(() => {});
}

/**
 * Fallback path: open the "Look up more records" dialog and pick a row.
 * You usually won't call this directly—selectLookupValue() calls it when needed.
 */
async function pickFromLookupDialog(
  page: Page,
  input: Locator,
  value: string,
  { timeout = 15000, allowStartsWith = true }: { timeout?: number; allowStartsWith?: boolean } = {}
) {
  // Try a nearby "Look up more records" button
  const container = input.locator('xpath=ancestor::*[self::div or self::section][@role="group" or contains(@class,"")][1]');
  const moreBtn = container.getByRole('button', { name: /look\s*up|more\s*records|search\s*for\s*more/i }).first();

  if (await moreBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
    await moreBtn.click();
  } else {
    // Keyboard gesture often opens the chooser
    await input.press('Alt+ArrowDown').catch(() => {});
  }

  // Dialog with a title like "Look Up Record(s)"
  const dlg = page.getByRole('dialog').filter({ hasText: /look\s*up/i }).first();
  await expect(dlg, 'Lookup dialog should open').toBeVisible({ timeout });

  // Search box (often "Search this view")
  const dlgSearch =
    dlg.getByRole('textbox', { name: /search/i }).first()
      .or(dlg.getByPlaceholder(/search/i));

  if (await dlgSearch.isVisible().catch(() => false)) {
    await dlgSearch.fill(value);
    await dlgSearch.press('Enter').catch(() => {});
  }

  // Grid rows (role="grid" or table); accessible name usually includes record name
  const grid = dlg.getByRole('grid').first().or(dlg.locator('[role="table"], [data-id*="grid"]'));
  const exactRx = new RegExp(`^${esc(value)}$`, 'i');
  const startsRx = new RegExp(`^${esc(value)}`, 'i');

  let row = grid.getByRole('row', { name: exactRx }).first();
  if (!(await row.isVisible({ timeout: 5000 }).catch(() => false)) && allowStartsWith) {
    row = grid.getByRole('row', { name: startsRx }).first();
  }
  await expect(row, `Row for "${value}" should appear in lookup grid`).toBeVisible({ timeout: 5000 });

  // Double-click to select; fall back to checkbox + Add/OK
  await row.dblclick().catch(async () => {
    const checkbox = row.getByRole('checkbox').first();
    if (await checkbox.isVisible().catch(() => false)) {
      await checkbox.check();
      await dlg.getByRole('button', { name: /add|ok|select/i }).click();
    } else {
      await row.click();
      await dlg.getByRole('button', { name: /add|ok|select/i }).click();
    }
  });

  await expect(dlg).toBeHidden({ timeout: 5000 });
}

/**
 * Main entry: select a value in a Model-Driven App lookup.
 * - Types into the lookup and tries to pick from the dropdown listbox.
 * - If not visible (virtualized deep down), falls back to the dialog.
 */
export async function selectLookupValue(
  page: Page,
  label: string,
  value: string,
  opts: {
    timeout?: number;
    exact?: boolean;
    dialogStartsWith?: boolean;
    scrollPages?: number;     // how many "pages" to attempt in the listbox before dialog
    openWithArrowDown?: boolean;
  } = {}
) {
  const {
    timeout = 15000,
    exact = true,
    dialogStartsWith = true,
    scrollPages = 8,
    openWithArrowDown = true,
  } = opts;

  const input = await findLookup(page, label);

  // Ensure visible & focused
  await input.scrollIntoViewIfNeeded().catch(() => {});
  await input.click({ timeout });

  // Clear any existing value
  await clearLookupValue(input);

  // Type the target value
  await input.fill(value, { timeout });

  // Open dropdown to reveal options
  if (openWithArrowDown) {
    await input.press('ArrowDown').catch(() => {});
  }

  // Wait for search/network to settle
  await waitForDataverseSearch(page, timeout);

  // Options usually live in a role="listbox" container
  const listboxes = page.getByRole('listbox');
  const listbox = (await listboxes.count()) > 0 ? listboxes.last() : input.locator('[role="listbox"]').last();

  // Build option locators (exact first, then startsWith for truncated names)
  const exactRx = new RegExp(`^${esc(value)}$`, 'i');
  const startsRx = new RegExp(`^${esc(value)}`, 'i');

  let option = listbox.getByRole('option', { name: exactRx }).first();

  // If visible quickly, click and return
  if (await option.isVisible({ timeout: 1000 }).catch(() => false)) {
    await option.click();
    return;
  }

  // Try startsWith (UI may ellipsis-truncate)
  if (!exact) {
    option = listbox.getByRole('option', { name: startsRx }).first();
    if (await option.isVisible({ timeout: 1000 }).catch(() => false)) {
      await option.click();
      return;
    }
  }

  // If still not visible, attempt to scroll the virtualized listbox a few "pages"
  for (let i = 0; i < scrollPages; i++) {
    // Scroll by the visible height to simulate a page down
    await listbox.evaluate((e: HTMLElement) => e.scrollBy(0, e.clientHeight));
    if (await option.isVisible({ timeout: 500 }).catch(() => false)) {
      await option.click();
      return;
    }
    if (!exact) {
      const opt2 = listbox.getByRole('option', { name: startsRx }).first();
      if (await opt2.isVisible({ timeout: 300 }).catch(() => false)) {
        await opt2.click();
        return;
      }
    }
  }

  // Escalate to dialog as the reliable fallback
  await pickFromLookupDialog(page, input, value, { timeout, allowStartsWith: dialogStartsWith });
}

/**
 * Optional utility: explicitly force the dialog path (useful for debugging or when you always want the grid).
 */
export async function forceLookupDialogSelection(
  page: Page,
  label: string,
  value: string,
  opts: { timeout?: number; dialogStartsWith?: boolean } = {}
) {
  const { timeout = 15000, dialogStartsWith = true } = opts;
  const input = await findLookup(page, label);
  await input.click();
  await clearLookupValue(input);
  await pickFromLookupDialog(page, input, value, { timeout, allowStartsWith: dialogStartsWith });
}
