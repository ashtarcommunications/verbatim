import { assert } from 'vitest';
import { browser, EXTENSION_ID } from '../setupTests';

describe('Verbatim Options', () => {
	it('Renders an options page', async () => {
		const page = await browser.newPage();
		await page.goto(`chrome-extension://${EXTENSION_ID}/options.html`);

		let enabled = await page.$eval('#enabled', (el) => el.checked);
		assert.isFalse(enabled, 'Toggle is off by default');
		await page.click('.switch');
		enabled = await page.$eval('#enabled', (el) => el.checked);
		assert.isTrue(enabled, 'Toggle is on');

		await page.type('#shortcut-copy', 'ctrl+1');
		const shortcutCopy = await page.$eval('#shortcut-copy', (el) => el.value);
		assert.strictEqual(
			shortcutCopy,
			'ctrl+1',
			'Shortcut copy is set correctly',
		);
	});
});
