const toCamelCase = (str) =>
	str.replace(/-([a-z])/g, (g) => g[1].toUpperCase());

const checkboxHandler = (e) => {
	browser.storage.sync.set({
		[toCamelCase(e.target.id)]: e.target.checked,
	});
};

const shortcutHandler = (e) => {
	browser.storage.sync.set({
		[toCamelCase(e.target.id)]: e.target.value
			.toLowerCase()
			.replace(' ', '')
			.trim(),
	});
};

document.addEventListener('DOMContentLoaded', async () => {
	const { version } = browser.runtime.getManifest();
	document.getElementById('version').innerText = `v${version}`;

	const settings = await browser.storage.sync.get(null);

	document.getElementById('enabled').checked = !!settings.enabled;
	document.getElementById('debug-mode').checked = !!settings.debugMode;

	document.getElementById('shortcut-tilde').value =
		settings.shortcutTilde || '';

	document
		.getElementById('enabled')
		.addEventListener('change', checkboxHandler);

	document
		.getElementById('debug-mode')
		.addEventListener('change', checkboxHandler);

	document
		.getElementById('shortcut-tilde')
		.addEventListener('input', shortcutHandler);
});
