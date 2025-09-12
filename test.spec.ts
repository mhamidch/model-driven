import { test } from '@playwright/test';
import {
  setText, setNumber, setBoolean, setOption, setMultiOption,
  setLookup, setDate, clickCommand, expectToast, getValue
} from '../libs/powerapps-uci-forms';

test('Create Contact (model-driven)', async ({ page }) => {
  await page.goto(process.env.POWERAPPS_APP_URL!);
  await page.getByRole('link', { name: /contacts/i }).click();
  await clickCommand(page, 'New');

  await setText(page, 'First Name', 'Test');
  await setText(page, 'Last Name', `Contact ${Date.now()}`);
  await setLookup(page, 'Account Name', 'Northwind Traders');
  await setOption(page, 'Preferred Contact Method', 'Email');
  await setBoolean(page, 'Do not allow Phone Calls', true);
  await setMultiOption(page, 'Topics', ['Billing', 'Onboarding']);
  await setNumber(page, 'Credit Limit', 12345.67);
  await setDate(page, 'Birthday', '23/10/1990');

  await clickCommand(page, 'Save & Close');
  await expectToast(page, /saved/i);

  // sanity readback
  const pcm = await getValue(page, 'Preferred Contact Method');
  test.expect(/email/i.test(pcm)).toBeTruthy();
});
