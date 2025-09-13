// libs/powerapps-uci-forms.ts
import { Page, Locator, expect } from '@playwright/test';

/* -------------------------- utilities & helpers -------------------------- */

const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const labelRx = (label: string) => new RegExp(`^${esc(label)}(\\s*\\*)?$`, 'i'); // allow required '*'

async function clearIfPossible(input: Locator) {
  try {
    await input.fill('');
  } catch {
    await input.click({ clickCount: 3 }).catch(() => {});
    await input.press('Backspace').catch(() => {});
  }
}

/** Find the first locator in the list that exists & is visible */
async function firstVisible(...locators: Locator[]): Promise<Locator | null> {
  for (const loc of locators) {
    if (await loc.count()) {
      if (await loc.first().isVisible().catch(() => false)) return loc.first();
    }
  }
  return null;
}

async function findTextbox(page: Page, label: string) {
  const tb = page.getByRole('textbox', { name: labelRx(label) }).first();
  await expect(tb, `Textbox "${label}" should exist`).toBeVisible();
  return tb;
}

async function findCombobox(page: Page, label: string) {
  // Covers normal fields and lookups (UCI appends ", Lookup")
  const cb = page.getByRole('combobox', { name: new RegExp(`^${esc(label)}(,\\s*Lookup)?(\\s*\\*)?$`, 'i') }).first();
  await expect(cb, `Combobox "${label}" should exist`).toBeVisible();
  return cb;
}

async function getListbox(page: Page, timeout = 10000) {
  const lb = page.getByRole('listbox').first();
  await expect(lb).toBeVisible({ timeout });
  return lb;
}

function monthYearText(d: Date): [string, string] {
  const long = d.toLocaleString('en-AU', { month: 'long', year: 'numeric' });
  const short = d.toLocaleString('en-AU', { month: 'short', year: 'numeric' });
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
  // Some “textarea” fields still expose role=textbox; fall back to aria-label textarea
  let area = page.getByRole('textbox', { name: labelRx(label) }).first();
  if (!(await area.count())) area = page.locator(`textarea[aria-label][aria-label*="${label}"]`).first();
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
  const digits = v.replace(/[^\d.-]/g, '');
  await expect(tb).toHaveValue(new RegExp(digits.replace('.', '\\.')));
}

/* ------------------------------- two-options ----------------------------- */

export async function setBoolean(page: Page, label: string, value: boolean) {
  const targetName = value ? /yes|true|on/i : /no|false|off/i;

  // 1) Combobox Yes/No
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

  // 2) Checkbox
  const maybeCheckbox = page.getByRole('checkbox', { name: labelRx(label) }).first();
  if (await maybeCheckbox.count()) {
    const checked = await maybeCheckbox.isChecked();
    if (checked !== value) await maybeCheckbox.click();
    await (value ? expect(maybeCheckbox).toBeChecked() : expect(maybeCheckbox).not.toBeChecked());
    return;
  }

  // 3) Switch
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
    const selected = await opt.getAttribute('aria-selected');
    if (selected !== 'true') await opt.click();
  }

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
  const nameRx = new RegExp(`^${esc(label)}(,\\s*Lookup)?(\\s*\\*)?$`, 'i');
  const lookup = page.getByRole('combobox', { name: nameRx }).first();
  await expect(lookup, `Lookup "${label}"`).toBeVisible();

  const clearBtn = page
    .locator('button[aria-label^="Clear"], button[title^="Clear"], button[aria-label^="Delete"]')
    .first();
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
    const more = lb.getByRole('option', { name: /search for more records|look up more records/i }).first();
    if (!(await more.count())) throw new Error(`No lookup result for "${value}" in "${label}".`);
    await more.click();

    const dlg = page.getByRole('dialog').first();
    await expect(dlg).toBeVisible();

    const search =
      (await firstVisible(
        dlg.getByRole('textbox', { name: /search/i }).first(),
        dlg.getByPlaceholder(/search this view/i).first()
      )) ?? dlg.locator('input[type="text"]').first();

    await search.fill(value);
    await search.press('Enter');

    const row = dlg.getByRole('row', { name: rx }).first();
    await expect(row).toBeVisible({ timeout: timeoutMs });
    await row.dblclick();
    await expect(dlg).toBeHidden();
  }

  const container =
    (await firstVisible(page.locator(`[aria-label="${label}, Lookup"]`).first(), lookup)) ?? lookup;
  await expect(container.getByText(rx)).toBeVisible({ timeout: timeoutMs });
}

export async function clearLookup(page: Page, label: string) {
  const container =
    (await firstVisible(
      page.locator(`[aria-label="${label}, Lookup"]`).first(),
      page.getByRole('combobox', { name: new RegExp(`^${esc(label)}(,\\s*Lookup)?(\\s*\\*)?$`, 'i') }).first()
    )) ?? page.locator('noop'); // will fail visibly if not found

  await expect(container).toBeVisible();

  const clearBtn = container
    .locator('button[aria-label^="Clear"], button[title^="Clear"], button[aria-label^="Delete"]')
    .first();
  if (await clearBtn.isVisible().catch(() => false)) await clearBtn.click();
}

/* ---------------------------------- date --------------------------------- */
/* Model-driven UCI Date: combobox-only (Out-of-the-box DatePicker) */

export type TargetDate = string | Date;

function parseDateLoose(input: TargetDate): Date {
  if (input instanceof Date) return input;
  const s = String(input).trim();

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const t = Date.parse(s);
    if (!Number.isNaN(t)) return new Date(t);
  }
  const m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/.exec(s);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  throw new Error(`Unrecognized date "${input}". Use dd/MM/yyyy or yyyy-MM-dd.`);
}

/** Open the UCI DatePicker callout/flyout for a given combobox label */
async function openDateCalendar(page: Page, combo: Locator): Promise<Locator> {
  // Try the explicit calendar button inside/near the field
  const btn = await firstVisible(
    combo.locator('button[aria-label*="calendar" i], button[title*="calendar" i]').first(),
    combo.locator('xpath=following::button[contains(@aria-label,"calendar")][1]').first()
  );

  if (btn) {
    await btn.click();
  } else {
    // Many UCI skins open calendar with Alt+ArrowDown
    await combo.click();
    await combo.press('Alt+ArrowDown').catch(() => {});
  }

  // The calendar usually appears in a callout or dialog; both contain .ms-DatePicker
  const cal = await firstVisible(
    page.locator('[role="dialog"] .ms-DatePicker, [role="dialog"] .DatePicker').first(),
    page.locator('.ms-DatePicker, .DatePicker').first()
  );

  if (!cal) throw new Error('Calendar UI did not open for the date field.');
  await expect(cal).toBeVisible({ timeout: 10000 });
  return cal;
}

async function navigateCalendar(cal: Locator, target: Date) {
  const header = cal.locator('[role="heading"], .ms-DatePicker-monthAndYear, .DatePicker-monthAndYear').first();
  await expect(header).toBeVisible({ timeout: 5000 });

  const next = cal.getByRole('button', { name: /next month|go to next month/i }).first();
  const prev = cal.getByRole('button', { name: /previous month|go to previous month/i }).first();

  const want = new Date(target.getFullYear(), target.getMonth(), 1);
  const [long, short] = monthYearText(target);

  const headerToMonth = async () => {
    const t = (await header.textContent())?.trim() ?? '';
    const m = /([A-Za-z]{3,})\s+(\d{4})/.exec(t);
    return m ? new Date(`${m[1]} 1, ${m[2]}`) : null;
  };

  const current = await headerToMonth();
  let maxHops = 480;
  let goNext: boolean | null = null;
  if (current) {
    const diff = (want.getFullYear() - current.getFullYear()) * 12 + (want.getMonth() - current.getMonth());
    maxHops = Math.abs(diff) + 24;
    goNext = diff > 0;
  }

  for (let i = 0; i < maxHops; i++) {
    const t = (await header.textContent())?.trim() ?? '';
    if (t.includes(long) || t.includes(short)) break;

    if (goNext === null) goNext = true;
    if (goNext && (await next.count())) await next.click();
    else if (await prev.count()) await prev.click();
    else await cal.press(goNext ? 'PageDown' : 'PageUp');
    await cal.waitFor({ state: 'visible' });
  }

  const dayBtn = cal.getByRole('button', { name: new RegExp(`^${target.getDate()}$`) }).first();
  await expect(dayBtn).toBeVisible({ timeout: 5000 });
  await dayBtn.click();
}

/** Combobox-only UCI date setter */
export async function setDate(page: Page, label: string, value: TargetDate) {
  const target = parseDateLoose(value);
  const combo = await findCombobox(page, label);

  // If there is an inner input, try typing first (some skins accept direct typing)
  const inner = (await firstVisible(
    combo.locator('input[type="text"]').first(),
    combo.locator('input').first()
  ));

  if (inner) {
    const dd = String(target.getDate()).padStart(2, '0');
    const mm = String(target.getMonth() + 1).padStart(2, '0');
    const yyyy = target.getFullYear();
    const au = `${dd}/${mm}/${yyyy}`;

    await inner.click();
    await inner.fill('');
    await inner.type(au, { delay: 20 });
    await inner.press('Enter').catch(() => {});
    await inner.blur().catch(() => {});

    // Accept AU or ISO rendering after commit
    const rx = new RegExp(`${esc(au)}|${esc(`${yyyy}-${mm}-${dd}`)}`);
    const val = await inner.evaluate((el: HTMLInputElement) => el.value).catch(() => '');
    if (rx.test(val)) return; // success by typing
  }

  // Fallback to picking from the calendar (reliable across month/year)
  const cal = await openDateCalendar(page, combo);
  await navigateCalendar(cal, target);

  // Verify value in input (if present) or in combobox text
  const dd = String(target.getDate()).padStart(2, '0');
  const mm = String(target.getMonth() + 1).padStart(2, '0');
  const yyyy = target.getFullYear();
  const rx = new RegExp(`${esc(`${dd}/${mm}/${yyyy}`)}|${esc(`${yyyy}-${mm}-${dd}`)}`);

  if (inner) {
    await expect(inner).toHaveValue(rx);
  } else {
    await expect(combo).toHaveText(rx);
  }
}

/* ------------------------- toolbar & notifications ----------------------- */

export async function clickCommand(page: Page, name: string) {
  const btn = page.getByRole('button', { name: new RegExp(`^${esc(name)}$`, 'i') }).first();
  await expect(btn, `Command "${name}"`).toBeVisible();
  await btn.click();
}

export async function expectToast(page: Page, rx: RegExp = /saved|success|created/i, timeout = 10000) {
  // No .or() — just try both patterns
  const toastA = page.getByRole('alert').filter({ hasText: rx }).first();
  const toastB = page.locator('[aria-live]').filter({ hasText: rx }).first();
  const toast = (await firstVisible(toastA, toastB)) ?? toastA;
  await expect(toast).toBeVisible({ timeout });
}

/* -------------------------------- readback -------------------------------- */

export async function getValue(page: Page, label: string): Promise<string> {
  const tb = page.getByRole('textbox', { name: labelRx(label) }).first();
  if (await tb.count()) return (await tb.inputValue().catch(async () => (await tb.textContent()) ?? '')) ?? '';

  const cb = page.getByRole('combobox', { name: new RegExp(`^${esc(label)}(,\\s*Lookup)?(\\s*\\*)?$`, 'i') }).first();
  if (await cb.count()) return (await cb.textContent())?.trim() ?? '';

  const cont = page.locator(`[aria-label="${label}"]`).first();
  if (await cont.count()) return (await cont.textContent())?.trim() ?? '';

  throw new Error(`Cannot read value for "${label}"`);
}
