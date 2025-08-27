import { assert } from 'vitest';
import { browser, EXTENSION_ID } from '../setupTests.js';

describe('Verbatim Popup', () => {
	it('Renders a popup', async () => {
		const page = await browser.newPage();
		await page.goto(`chrome-extension://${EXTENSION_ID}/popup.html`);

		let enabled = await page.$eval('#switch', (el) => el.checked);
		assert.isFalse(enabled, 'Toggle is off by default');
		await page.click('.switch');
		enabled = await page.$eval('#switch', (el) => el.checked);
		assert.isTrue(enabled, 'Toggle is on');

		const href = await page.$eval('#options-link', (el) =>
			el.getAttribute('href'),
		);
		assert.strictEqual(
			href,
			`chrome-extension://${EXTENSION_ID}/options.html`,
			'Correct options link',
		);
	});
});
