const onOffHandler = async () => {
	browser.storage.sync.set({
		enabled: document.getElementById('switch').checked,
	});
};

document.addEventListener('DOMContentLoaded', async () => {
	const settings = await browser.storage.sync.get('enabled');

	document.getElementById('switch').checked = !!settings.enabled;
	document.getElementById('switch').addEventListener('change', onOffHandler);

	document.getElementById('options-link').href =
		browser.runtime.getURL('options.html');
	document.getElementById('options-link').target = '_blank';
});
