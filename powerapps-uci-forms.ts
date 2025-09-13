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
// --- replace your existing setDate + helpers with this block ---

import { Page, Locator, expect } from '@playwright/test';

const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const labelRx = (label: string) => new RegExp(`^${esc(label)}(\\s*\\*)?$`, 'i'); // allow required asterisk

type TargetDate = string | Date;

/** Parse "dd/MM/yyyy", "yyyy-MM-dd", or Date into a Date */
function parseDateLoose(input: TargetDate): Date {
  if (input instanceof Date) return input;
  const s = input.trim();

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const t = Date.parse(s);
    if (!Number.isNaN(t)) return new Date(t);
  }
  const m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/.exec(s);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  throw new Error(`Unrecognized date "${input}". Use dd/MM/yyyy or yyyy-MM-dd.`);
}

/** Try to locate an editable input for a date field OR the combobox container */
async function findDateHandles(page: Page, label: string): Promise<{
  editableInput?: Locator;   // if present, we can type/verify here
  comboOrText: Locator;      // main control (textbox or combobox)
  openButton: Locator;       // button that opens the calendar
}> {
  // 1) Prefer a visible textbox by accessible name (many tenants still use this)
  let tb = page.getByRole('textbox', { name: labelRx(label) }).first();

  if (await tb.isVisible().catch(() => false)) {
    // calendar button is usually adjacent or a generic "Calendar" button
    let openBtn = tb.locator('xpath=following::button[contains(@aria-label,"calendar")][1]').first();
    if (!(await openBtn.count())) {
      openBtn = page.getByRole('button', { name: /open.*calendar|calendar|show calendar/i }).first();
    }
    return { editableInput: tb, comboOrText: tb, openButton: openBtn };
  }

  // 2) Combobox by accessible label (common for date fields in UCI)
  const comboByRole = page.getByRole('combobox', { name: labelRx(label) }).first();
  if (await comboByRole.isVisible().catch(() => false)) {
    // Some combos have an inner input; grab it if present
    const innerInput = comboByRole.locator('input[type="text"], input[role="spinbutton"], input').first();
    const editable = (await innerInput.count()) ? innerInput : undefined;

    // Find calendar button near the combobox
    let openBtn = comboByRole.locator('button[aria-label*="calendar" i], button[title*="calendar" i]').first();
    if (!(await openBtn.count())) {
      // try a nearby button in the same field container
      openBtn = comboByRole.locator('xpath=following::button[contains(@aria-label,"calendar")][1]').first()
        .or(page.getByRole('button', { name: /open.*calendar|calendar|show calendar/i }).first());
    }

    return { editableInput: editable, comboOrText: comboByRole, openButton: openBtn };
  }

  // 3) Last chance: aria-label container + descendants
  const container = page.locator(`[aria-label="${label}"], [aria-label^="${label},"]`).first();
  if (await container.isVisible().catch(() => false)) {
    const maybeInput = container.locator('input').first();
    const openBtn = container.locator('button[aria-label*="calendar" i], button[title*="calendar" i]').first()
      .or(page.getByRole('button', { name: /open.*calendar|calendar|show calendar/i }).first());
    return {
      editableInput: (await maybeInput.count()) ? maybeInput : undefined,
      comboOrText: (await maybeInput.count()) ? maybeInput : container,
      openButton: openBtn
    };
  }

  throw new Error(`Could not find date control for label "${label}". It may be renamed or inside a dialog.`);
}

/** Navigate the calendar popup to the desired month/year and click the day */
async function pickFromCalendar(page: Page, target: Date) {
  // Calendar is usually a dialog containing a grid
  const cal = page.getByRole('dialog').filter({ has: page.locator('[role="grid"]') }).first()
    .or(page.locator('.ms-DatePicker, .DatePicker').first());
  await expect(cal).toBeVisible({ timeout: 10_000 });

  const header = cal.locator('[role="heading"], .ms-DatePicker-monthAndYear, .DatePicker-monthAndYear').first();
  const next = cal.getByRole('button', { name: /next month|go to next month/i }).first();
  const prev = cal.getByRole('button', { name: /previous month|go to previous month/i }).first();

  const want = new Date(target.getFullYear(), target.getMonth(), 1);
  const long = target.toLocaleString('en-AU', { month: 'long', year: 'numeric' });
  const short = target.toLocaleString('en-AU', { month: 'short', year: 'numeric' });

  // Try bounded hops
  const headerToMonth = async () => {
    const t = (await header.textContent())?.trim() ?? '';
    const m = /([A-Za-z]{3,})\s+(\d{4})/.exec(t);
    return m ? new Date(`${m[1]} 1, ${m[2]}`) : null;
  };
  const current = await headerToMonth();
  let maxHops = 480, goNext: boolean | null = null;
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

  const dayBtn = cal.getByRole('button', { name: new RegExp(`^${target.getDate()}$`) }).first()
    .or(cal.locator('[role="gridcell"]').filter({ hasText: String(target.getDate()) }).first());
  await expect(dayBtn).toBeVisible();
  await dayBtn.click();
}

/** Public: set a date in UCI whether it’s textbox-style or combobox-style */
export async function setDate(page: Page, label: string, value: TargetDate) {
  const target = parseDateLoose(value);
  const { editableInput, comboOrText, openButton } = await findDateHandles(page, label);

  // If there is an editable input, first try a clean type+commit path
  if (editableInput) {
    await editableInput.click();
    // clear existing
    await editableInput.fill('');
    // AU format typing (UCI usually accepts typed values)
    const dd = String(target.getDate()).padStart(2, '0');
    const mm = String(target.getMonth() + 1).padStart(2, '0');
    const yyyy = target.getFullYear();
    const au = `${dd}/${mm}/${yyyy}`;

    await editableInput.type(au, { delay: 20 });
    // Commit value (Enter or Tab often triggers validation)
    await editableInput.press('Enter').catch(() => {});
    await editableInput.blur().catch(() => {});

    // Verify; if it didn’t stick, fall back to calendar selection
    const valueOk = await editableInput.evaluate((el: HTMLInputElement) => el.value).catch(() => '');
    if (valueOk && (new RegExp(`${esc(au)}|${yyyy}-${mm}-${dd}`)).test(valueOk)) {
      return;
    }
  }

  // Either there was no editable input, or typing didn’t commit correctly -> use calendar
  await comboOrText.click().catch(() => {});
  if (openButton && (await openButton.isVisible().catch(() => false))) {
    await openButton.click();
  } else {
    // keyboard open (Alt+ArrowDown often works)
    await comboOrText.press('Alt+ArrowDown').catch(() => {});
  }
  await pickFromCalendar(page, target);

  // Final verification (read from whichever input we have; else check combobox text)
  const dd = String(target.getDate()).padStart(2, '0');
  const mm = String(target.getMonth() + 1).padStart(2, '0');
  const yyyy = target.getFullYear();
  const rx = new RegExp(`${esc(`${dd}/${mm}/${yyyy}`)}|${esc(`${yyyy}-${mm}-${dd}`)}`);

  if (editableInput) {
    await expect(editableInput).toHaveValue(rx);
  } else {
    // Some combobox variants render the value as text content
    await expect(comboOrText).toHaveText(rx);
  }
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
