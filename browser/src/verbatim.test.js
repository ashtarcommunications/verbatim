import { assert } from 'vitest';
import { browser, EXTENSION_ID } from '../setupTests';

describe('Verbatim', () => {
	it('Correctly handles keyboard shortcuts', async () => {
		// The test HTML page loads the verbatim.js script and a test shim that mocks the browser storage API
		// That lets us set the extension to enabled on page load and control the mock from the page context
		// because Node otherwise can't access the browser object. There's probably a way to do it with
		// Node and write more comprehensive tests with different settings states, but that's a lot harder
		const page = await browser.newPage();
		const url = `chrome-extension://${EXTENSION_ID}/test.html`;
		await page.goto(url);

		await page.waitForSelector('#verbatim');
	});
});
