import { test, expect } from '@playwright/test';

test('page loads and has content', async ({ page }) => {
  const response = await page.goto('/');
  console.log('Status:', response?.status());

  // Take screenshot of what we see
  await page.screenshot({ path: 'test-results/smoke-screenshot.png' });

  // Check page title
  const title = await page.title();
  console.log('Title:', title);

  // Check if body has content
  const bodyText = await page.locator('body').textContent();
  console.log('Body length:', bodyText?.length);
  console.log('Body preview:', bodyText?.substring(0, 200));

  // Check for our elements
  const hasDropZone = await page.locator('.drop-zone').count();
  console.log('Drop zone count:', hasDropZone);

  const hasDarkMode = await page.locator('#darkModeLanding').count();
  console.log('Dark mode button count:', hasDarkMode);

  expect(title).toContain('Oxi');
});
