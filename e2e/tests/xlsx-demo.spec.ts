import { test, expect } from '@playwright/test';

const PAGE = 'http://localhost:8080/web/xlsx-demo.html';

/// Read the sheet canvas: how big it is and how much is drawn on it.
async function drawn(page: import('@playwright/test').Page) {
  return page.locator('#sheetInk').evaluate((held: HTMLCanvasElement) => {
    const pen = held.getContext('2d');
    const { data } = pen!.getImageData(0, 0, held.width, held.height);
    let ink = 0;
    for (let at = 0; at < data.length; at += 4) {
      if (data[at] < 250 || data[at + 1] < 250 || data[at + 2] < 250) ink += 1;
    }
    return { width: held.width, height: held.height, ink };
  });
}

/// Wait for the workbook to be open and drawn. Until then the page is a shell
/// with the right names on it and no sheet behind them.
async function sheetIsUp(page: import('@playwright/test').Page) {
  await expect(page.locator('#state')).toBeHidden({ timeout: 30_000 });
  await expect(page.locator('#tabs button').first()).toBeVisible();
}

/// A point inside the grid, past the strip of column letters and the column of
/// row numbers.
async function onTheGrid(page: import('@playwright/test').Page) {
  return page.evaluate(() => {
    const box = document.getElementById('sheet')!.getBoundingClientRect();
    return { x: box.left + 200, y: box.top + 80 };
  });
}

/// Walk the cursor to C5 of the sample sheet, one key at a time, waiting for
/// each step to land. Pressing four arrows in a row and looking only at where
/// it stopped cannot tell "the keys were dropped" from "the cursor moved
/// wrongly", and one of them is a bug.
async function walkToC5(page: import('@playwright/test').Page) {
  // From A1, where a freshly opened sheet leaves the cursor. Waiting for the
  // name box to READ A1 proves nothing — the HTML says A1 before anything has
  // loaded, and an arrow key at that moment does nothing because there is no
  // sheet yet. What has to be waited for is the sheet arriving.
  await sheetIsUp(page);
  await expect(page.locator('#where')).toHaveText('A1');
  for (const seat of ['A2', 'A3', 'A4', 'A5']) {
    await page.keyboard.press('ArrowDown');
    await expect(page.locator('#where')).toHaveText(seat);
  }
  for (const seat of ['B5', 'C5']) {
    await page.keyboard.press('ArrowRight');
    await expect(page.locator('#where')).toHaveText(seat);
  }
}

/// The sheet is painted, not laid out in DOM cells, so the only proof that the
/// page works is that there is ink on the canvas — and that nothing complained
/// on the way there.
test('the sheet demo draws the sample workbook', async ({ page }) => {
  const complaints: string[] = [];
  page.on('console', message => {
    if (message.type() === 'error') complaints.push(message.text());
  });
  page.on('pageerror', error => complaints.push(String(error)));

  await page.goto(PAGE);
  await page.locator('#sheetInk').waitFor({ timeout: 30_000 });
  await sheetIsUp(page);

  const shown = await drawn(page);
  console.log('canvas', shown);
  expect(shown.width).toBeGreaterThan(100);
  expect(shown.height).toBeGreaterThan(100);
  expect(shown.ink).toBeGreaterThan(500);
  expect(complaints).toEqual([]);

  // A canvas has no cells to click, so every gesture goes through one hit
  // test. Clicking inside the grid has to land the cursor on the cell under
  // the pointer rather than leaving it where it was.
  const seat = await onTheGrid(page);
  await page.mouse.click(seat.x, seat.y);
  await expect(page.locator('#where')).not.toHaveText('A1');
  await expect(page.locator('#where')).toHaveText(/^[A-Z]+\d+$/);

  // The panel beside it says what the sheet is made of.
  const facts = await page.locator('#facts').textContent();
  expect(facts).toContain('Standard font');

  await page.screenshot({ path: 'test-results/xlsx-demo.png', fullPage: false });
});

/// Every element answers to a name of its own. Two of them sharing an id is
/// not a tidiness complaint: getElementById returns whichever comes first, so
/// the later one is unreachable and whatever wanted it silently gets the other
/// — which is how the canvas and the colour picker both being `ink` left the
/// sheet unpainted while every unit test stayed green.
test('no two elements answer to the same id', async ({ page }) => {
  await page.goto(PAGE);
  const twice = await page.evaluate(() => {
    const seen = new Set<string>();
    const doubled: string[] = [];
    for (const one of document.querySelectorAll('[id]')) {
      if (seen.has(one.id)) doubled.push(one.id);
      seen.add(one.id);
    }
    return doubled;
  });
  expect(twice).toEqual([]);
});

/// Typing over a cell, the way it is done in Excel: you do not open the cell
/// first, the first key starts the value. Then the formula has to work itself
/// out, and undo has to put back what was there.
test('a cell can be typed over, worked out, and put back', async ({ page }) => {
  await page.goto(PAGE);
  await walkToC5(page);
  const was = await page.locator('#formula').inputValue();

  await page.keyboard.type('12345');
  await expect(page.locator('#entry')).toBeVisible();
  await page.keyboard.press('Enter');
  await expect(page.locator('#where')).toHaveText('C6');
  await page.keyboard.press('ArrowUp');
  await expect(page.locator('#formula')).toHaveValue('12345');
  await expect(page.locator('#save')).toBeEnabled();

  await page.locator('#back').click();
  await expect(page.locator('#where')).toHaveText('C5');
  await expect(page.locator('#formula')).toHaveValue(was!);
});

/// Freezing holds everything above and left of the cursor. The sheet is
/// painted in bands to do it, so what proves it is that the canvas changes and
/// the control says the freeze is on.
test('the panes can be frozen and let go', async ({ page }) => {
  await page.goto(PAGE);
  await walkToC5(page);
  const before = (await drawn(page)).ink;

  await page.locator('#freeze').click();
  await expect(page.locator('#freeze')).toHaveText('Unfreeze');
  expect((await drawn(page)).ink).toBeGreaterThan(0);

  await page.locator('#freeze').click();
  await expect(page.locator('#freeze')).toHaveText('Freeze');
  expect((await drawn(page)).ink).toBeGreaterThan(before / 2);
});
