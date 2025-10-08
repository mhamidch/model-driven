async function pickFromLookupDialog(
  page: Page,
  input: Locator,
  value: string,
  { timeout = 15000 }: { timeout?: number } = {}
) {
  // The “Look up more records” button is adjacent to the input.
  // It can be exposed as a button with an accessible name like:
  // "Look up more records", "Search for more records", or “Look up records”.
  const moreBtn = input
    .locator('xpath=ancestor::*[contains(@class,"quickCreate") or contains(@role,"group") or contains(@data-id,"lookup")]')
    .getByRole('button', { name: /look\s*up|more\s*records|search\s*for\s*more/i })
    .first();

  if (await moreBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
    await moreBtn.click();
  } else {
    // Some skins open the dialog via Alt+Down then End; emulate a user gesture
    await input.press('Alt+ArrowDown').catch(() => {});
    // If that didn’t open a dialog, try clicking possible “more” icon near the field
    // as a last resort (icon buttons may have only an aria-label tooltip).
  }

  // Dialog usually has role="dialog" and a title like "Look Up Record(s)"
  const dlg = page.getByRole('dialog').filter({ hasText: /look\s*up/i }).first();
  await expect(dlg, 'Lookup dialog should open').toBeVisible({ timeout });

  // Search box often appears as role="textbox" with placeholder "Search this view"
  const dlgSearch = dlg.getByRole('textbox', { name: /search/i }).first().or(dlg.getByPlaceholder(/search/i));
  if (await dlgSearch.isVisible().catch(() => false)) {
    await dlgSearch.fill(value);
    await dlgSearch.press('Enter').catch(() => {});
  }

  // Wait for grid rows to load; rows are often role="row" with cells role="gridcell"
  const grid = dlg.getByRole('grid').first().or(dlg.locator('[role="table"], [data-id*="grid"]'));
  const rowExact = grid.getByRole('row', { name: new RegExp(`^${esc(value)}$`, 'i') }).first();
  const rowStarts = grid.getByRole('row', { name: new RegExp(`^${esc(value)}`, 'i') }).first();

  const row = (await rowExact.isVisible({ timeout: 5000 }).catch(() => false)) ? rowExact : rowStarts;
  await expect(row, `Row for "${value}" should appear in lookup grid`).toBeVisible({ timeout: 5000 });

  // Double-click row selects it; alternatively tick checkbox then click Add
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

  // Ensure dialog closes and value is committed
  await expect(dlg).toBeHidden({ timeout: 5000 });
}
