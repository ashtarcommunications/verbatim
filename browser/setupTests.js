import puppeteer from 'puppeteer';

export const EXTENSION_PATH = 'src/';
export const EXTENSION_ID = 'xxxxxxxxxxxxxxxxxxxxxxx';

export let browser;

beforeEach(async () => {
	browser = await puppeteer.launch({
		headless: true,
		args: [
			`--disable-extensions-except=${EXTENSION_PATH}`,
			`--load-extension=${EXTENSION_PATH}`,
			`--no-sandbox`,
		],
	});
});

afterEach(async () => {
	await browser.close();
	browser = undefined;
});

export default null;
