// libs/powerapps-uci-forms.ts
import { Page, Locator, expect } from '@playwright/test';

/* -------------------------- utilities & helpers -------------------------- */

const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

async function clearIfPossible(input: Locator) {
  try {
    await input.fill('');
  } catch {
    // some UCI controls ignore fill(''); fall back to triple-click + Backspace
    await input.click({ clickCount: 3 }).catch(() => {});
    await input.press('Backspace').catch(() => {});
  }
}

function labelRx(label: string) {
  return new RegExp(`^${esc(label)}$`, 'i');
}

async function findTextbox(page: Page, label: string) {
  // UCI usually exposes text fields as role textbox with the label as accessible name
  const tb = page.getByRole('textbox', { name: labelRx(label) }).first();
  await expect(tb, `Textbox "${label}" should exist`).toBeVisible();
  return tb;
}

async function findCombobox(page: Page, label: string) {
  // Option sets and lookups are ARIA combobox
  const cb = page.getByRole('combobox', { name: new RegExp(`^${esc(label)}(,\\s*Lookup)?$`, 'i') }).first();
  await expect(cb, `Combobox "${label}" should exist`).toBeVisible();
  return cb;
}

async function getListbox(page: Page, timeout = 10000) {
  const lb = page.getByRole('listbox').first();
  await expect(lb).toBeVisible({ timeout });
  return lb;
}

function monthYearText(d: Date): [string, string] {
  const long = d.toLocaleString('en-AU', { month: 'long', year: 'numeric' }); // "September 2025"
  const short = d.toLocaleString('en-AU', { month: 'short', year: 'numeric' }); // "Sep 2025"
  return [long, short];
}

/* --------------------------------- text ---------------------------------- */

export async function setText(page: Page, label: string, value: string) {
  const tb = await findTextbox(page, label);
  await clearIfPossible(tb);
  await tb.type(value, { delay: 10 });
  await tb.blur();
  await expect(tb).toHaveValue(new RegExp(`^${esc(value)}$`));
}

export async function setTextArea(page: Page, label: string, value: string) {
  // Same locator pattern; some text areas expose role=textarea, fall back if needed
  let area = page.getByRole('textbox', { name: labelRx(label) }).first();
  if (!(await area.count())) area = page.locator('textarea[aria-label]').filter({ hasText: label }).first();
  await expect(area, `Textarea "${label}"`).toBeVisible();
  await clearIfPossible(area);
  await area.type(value, { delay: 10 });
  await area.blur();
  await expect(area).toHaveValue(new RegExp(esc(value)));
}

/* -------------------------------- number --------------------------------- */

export async function setNumber(page: Page, label: string, value: number | string) {
  const v = String(value);
  const tb = await findTextbox(page, label);
  await clearIfPossible(tb);
  await tb.type(v, { delay: 10 });
  await tb.blur();
  // Accept minor formatting (commas) by checking the digits
  const digits = v.replace(/[^\d.-]/g, '');
  await expect(tb).toHaveValue(new RegExp(digits.replace('.', '\\.')));
}

/* ------------------------------- two-options ----------------------------- */

export async function setBoolean(page: Page, label: string, value: boolean) {
  // In UCI, two-options can be a switch, checkbox, or combobox with Yes/No
  const targetName = value ? /yes|true|on/i : /no|false|off/i;

  // 1) Try combobox Yes/No
  const maybeCombo = page.getByRole('combobox', { name: labelRx(label) }).first();
  if (await maybeCombo.count()) {
    await maybeCombo.click();
    const lb = await getListbox(page);
    const opt = lb.getByRole('option', { name: targetName }).first();
    await expect(opt).toBeVisible();
    await opt.click();
    await expect(maybeCombo).toHaveText(targetName);
    return;
  }

  // 2) Try checkbox
  const maybeCheckbox = page.getByRole('checkbox', { name: labelRx(label) }).first();
  if (await maybeCheckbox.count()) {
    const checked = await maybeCheckbox.isChecked();
    if (checked !== value) await maybeCheckbox.click();
    await expect(maybeCheckbox)[value ? 'toBeChecked' : 'not.toBeChecked']();
    return;
  }

  // 3) Try switch (role="switch")
  const maybeSwitch = page.getByRole('switch', { name: labelRx(label) }).first();
  if (await maybeSwitch.count()) {
    const ariaChecked = (await maybeSwitch.getAttribute('aria-checked')) === 'true';
    if (ariaChecked !== value) await maybeSwitch.click();
    await expect(maybeSwitch).toHaveAttribute('aria-checked', value ? 'true' : 'false');
    return;
  }

  throw new Error(`Boolean control "${label}" not found as combobox/checkbox/switch.`);
}

/* ------------------------------ option sets ------------------------------ */

export async function setOption(page: Page, label: string, value: string, { exact = true } = {}) {
  const cb = await findCombobox(page, label);
  await cb.click();
  const lb = await getListbox(page);
  const rx = exact ? new RegExp(`^${esc(value)}$`, 'i') : new RegExp(value, 'i');
  const opt = lb.getByRole('option', { name: rx }).first();
  await expect(opt, `Option "${value}" for "${label}"`).toBeVisible();
  await opt.click();
  await expect(cb).toHaveText(rx);
}

export async function setMultiOption(page: Page, label: string, values: string[]) {
  const cb = await findCombobox(page, label);
  await cb.click();
  const lb = await getListbox(page);

  for (const v of values) {
    const rx = new RegExp(`^${esc(v)}$`, 'i');
    const opt = lb.getByRole('option', { name: rx }).first();
    await expect(opt, `Multi option "${v}"`).toBeVisible();
    // Ensure it's selected; UCI options often toggle selected state
    const selected = await opt.getAttribute('aria-selected');
    if (selected !== 'true') await opt.click();
  }

  // Close dropdown (Esc) and verify tags/chips show all selections
  await cb.press('Escape').catch(() => {});
  for (const v of values) {
    await expect(page.getByText(new RegExp(`^${esc(v)}$`, 'i'))).toBeVisible();
  }
}

/* -------------------------------- lookups -------------------------------- */

export async function setLookup(
  page: Page,
  label: string,
  value: string,
  { exact = true, timeoutMs = 12_000 }: { exact?: boolean; timeoutMs?: number } = {}
) {
  const nameRx = new RegExp(`^${esc(label)}(,\\s*Lookup)?$`, 'i');
  const lookup = page.getByRole('combobox', { name: nameRx }).first();
  await expect(lookup, `Lookup "${label}"`).toBeVisible();

  // Clear existing chip (if any)
  const clearBtn = page.locator(
    'button[aria-label^="Clear"], button[title^="Clear"], button[aria-label^="Delete"]'
  ).first();
  if (await clearBtn.isVisible().catch(() => false)) await clearBtn.click();

  await lookup.click();
  await lookup.fill('');
  await lookup.type(value, { delay: 20 });
  await lookup.press('ArrowDown').catch(() => {});

  const lb = await getListbox(page, timeoutMs);
  const rx = exact ? new RegExp(`^${esc(value)}$`, 'i') : new RegExp(value, 'i');
  let opt = lb.getByRole('option', { name: rx }).first();

  if (await opt.count()) {
    await opt.click();
  } else {
    // Fallback: "Search/Look up for more records"
    const more = lb
      .getByRole('option', { name: /search for more records|look up more records/i })
      .first();
    if (!(await more.count())) throw new Error(`No lookup result for "${value}" in "${label}".`);
    await more.click();

    const dlg = page.getByRole('dialog').first();
    await expect(dlg).toBeVisible();

    const search = dlg.getByRole('textbox', { name: /search/i }).first()
      .or(dlg.getByPlaceholder(/search this view/i).first());
    await search.fill(value);
    await search.press('Enter');

    const row = dlg.getByRole('row', { name: rx }).first();
    await expect(row).toBeVisible({ timeout: timeoutMs });
    // Double-click selects and closes
    await row.dblclick();
    await expect(dlg).toBeHidden();
  }

  // Verify pill set
  const container = page.locator(`[aria-label="${label}, Lookup"]`).first().or(lookup);
  await expect(container.getByText(rx)).toBeVisible({ timeout: timeoutMs });
}

export async function clearLookup(page: Page, label: string) {
  const nameRx = new RegExp(`^${esc(label)}(,\\s*Lookup)?$`, 'i');
  const container = page.locator(`[aria-label="${label}, Lookup"]`).first()
    .or(page.getByRole('combobox', { name: nameRx }).first());
  await expect(container).toBeVisible();

  const clearBtn = container.locator(
    'button[aria-label^="Clear"], button[title^="Clear"], button[aria-label^="Delete"]'
  ).first();
  if (await clearBtn.isVisible()) await clearBtn.click();
  // No chip with that label should remain
  await expect(container.getByText(new RegExp(esc(label), 'i'))).not.toBeVisible({ timeout: 2000 }).catch(() => {});
}

/* ---------------------------------- date --------------------------------- */

export type TargetDate = string | Date; // accepts "dd/MM/yyyy", "yyyy-MM-dd", or Date

function parseDateLoose(input: TargetDate): Date {
  if (input instanceof Date) return input;
  const s = input.trim();

  // Try ISO yyyy-MM-dd
  const dIso = Date.parse(s);
  if (!Number.isNaN(dIso) && /^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(dIso);

  // Try dd/MM/yyyy
  const m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/.exec(s);
  if (m) {
    const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
    return new Date(yyyy, mm - 1, dd);
  }
  throw new Error(`Unrecognized date "${input}". Use dd/MM/yyyy, yyyy-MM-dd, or Date.`);
}

export async function setDate(page: Page, label: string, value: TargetDate) {
  const target = parseDateLoose(value);

  // Many UCI date fields are textboxes with an adjacent "Open the calendar" button
  const tb = await findTextbox(page, label);
  await tb.click();

  // Open calendar via adjacent button or keyboard
  let openBtn = tb.locator('xpath=following::button[contains(@aria-label,"calendar")][1]').first();
  if (!(await openBtn.count())) {
    openBtn = page.getByRole('button', { name: /open.*calendar|calendar|show calendar/i }).first();
  }
  await openBtn.click();

  const cal = page.getByRole('dialog').filter({ has: page.locator('[role="grid"]') }).first()
             .or(page.locator('.ms-DatePicker, .DatePicker').first());
  await expect(cal).toBeVisible();

  const header = cal.locator('[role="heading"], .ms-DatePicker-monthAndYear, .DatePicker-monthAndYear').first();
  const next = cal.getByRole('button', { name: /next month|go to next month/i }).first();
  const prev = cal.getByRole('button', { name: /previous month|go to previous month/i }).first();

  const want = new Date(target.getFullYear(), target.getMonth(), 1);
  const [long, short] = monthYearText(target);

  // Try to parse current header month to compute bounded hops
  const headerToMonth = async () => {
    const t = (await header.textContent())?.trim() ?? '';
    const m = /([A-Za-z]{3,})\s+(\d{4})/.exec(t);
    return m ? new Date(`${m[1]} 1, ${m[2]}`) : null;
  };

  const current = await headerToMonth();
  let maxHops = 480; // default ~40y
  let goNext: boolean | null = null;
  if (current) {
    const diff = (want.getFullYear() - current.getFullYear()) * 12 + (want.getMonth() - current.getMonth());
    maxHops = Math.abs(diff) + 24; // 2y buffer
    goNext = diff > 0;
  }

  for (let i = 0; i < maxHops; i++) {
    const t = (await header.textContent())?.trim() ?? '';
    if (t.includes(long) || t.includes(short)) break;

    if (goNext === null) goNext = true; // optimistic
    if (goNext && (await next.count())) await next.click();
    else if (await prev.count()) await prev.click();
    else await cal.press(goNext ? 'PageDown' : 'PageUp');

    await cal.waitFor({ state: 'visible' });
  }

  // Click the day
  const dayBtn = cal.getByRole('button', { name: new RegExp(`^${target.getDate()}$`) }).first()
                .or(cal.locator('[role="gridcell"]').filter({ hasText: String(target.getDate()) }).first());
  await expect(dayBtn).toBeVisible();
  await dayBtn.click();

  // Verify textbox value (accept AU or ISO display)
  const dd = String(target.getDate()).padStart(2, '0');
  const mm = String(target.getMonth() + 1).padStart(2, '0');
  const yyyy = target.getFullYear();
  const au = `${dd}/${mm}/${yyyy}`;
  const iso = `${yyyy}-${mm}-${dd}`;
  await expect(tb).toHaveValue(new RegExp(`${esc(au)}|${esc(iso)}`));
}

/* ------------------------- toolbar & notifications ----------------------- */

export async function clickCommand(page: Page, name: string) {
  // Works for main ribbon/command bar buttons like "Save", "New", "Deactivate"
  const btn = page.getByRole('button', { name: new RegExp(`^${esc(name)}$`, 'i') }).first();
  await expect(btn, `Command "${name}"`).toBeVisible();
  await btn.click();
}

export async function expectToast(page: Page, rx: RegExp = /saved|success|created/i, timeout = 10000) {
  // UCI toasts are role="alert" or have aria-live
  const toast = page.getByRole('alert').filter({ hasText: rx }).first()
    .or(page.locator('[aria-live]').filter({ hasText: rx }).first());
  await expect(toast).toBeVisible({ timeout });
}

/* -------------------------------- readback -------------------------------- */

export async function getValue(page: Page, label: string): Promise<string> {
  // Try textbox first
  const tb = page.getByRole('textbox', { name: labelRx(label) }).first();
  if (await tb.count()) {
    return (await tb.inputValue().catch(async () => (await tb.textContent()) ?? '')) ?? '';
  }
  // Try combobox text content
  const cb = page.getByRole('combobox', { name: new RegExp(`^${esc(label)}(,\\s*Lookup)?$`, 'i') }).first();
  if (await cb.count()) return (await cb.textContent())?.trim() ?? '';
  // Fallback: read by aria-label container
  const cont = page.locator(`[aria-label="${label}"]`).first();
  if (await cont.count()) return (await cont.textContent())?.trim() ?? '';
  throw new Error(`Cannot read value for "${label}"`);
}
