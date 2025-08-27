var sendButton;
var innerFrame;

// Save settings for later use
var objSettings = new Object();

function setSettings(callback) {
	chrome.storage.sync.get(null, function (settings) {
		objSettings.enabled = settings.enabled;

		callback();
	});
}

// https://gist.github.com/WiliTest/b78ad2c234565ba8ce40df15440540d9
var editingIFrame = document.getElementsByClassName(
	'docs-texteventtarget-iframe',
)[0];
// if (editingIFrame) {
editingIFrame.contentDocument.addEventListener('keydown', hook, false);
// }

async function hook(e) {
	var keyCode = e.keyCode;
	console.log('keycode: ' + keyCode);
	if (keyCode === 192) {
		e.preventDefault();
		console.log('tilde pressed');

		console.log('sending message to inner frame');
		// console.log(innerFrame);
		innerFrame.postMessage('sendToSpeech', '*');
		// var outer = document.querySelector('.script-application-sidebar-content iframe').contentDocument;
		// var sandbox = outer.getElementById('sandboxFrame').contentDocument;
		// var inner = sandbox.getElementById('userHtmlFrame').contentDocument
		// inner.getElementById('send').click();
		//sendButton.click();
		// document.execCommand('copy');

		// chrome.runtime.sendMessage({ action: 'sendToSpeech' }, function (response) {
		//     console.log(response);
		// });
	}
}

window.addEventListener(
	'message',
	(event) => {
		if (event.data === 'innerFrame') {
			console.log('received inner frame');
			innerFrame = event.source;
		}
	},
	false,
);

// Keyboard shortcut functions
// Mousetrap.bind('`', function(e) {
//     console.log('Moustrap ran');
//     alert('You pressed tilde');

//     if (window.getSelection) {
//         //Send message to the background page with text to copy - tack selection onto cite
//         chrome.extension.sendMessage({ "text" : window.getSelection });
//     }

//     return false;
// });

window.addEventListener('load', function () {
	console.log('All assets are loaded');
	// document.querySelector('.goog-menu-item').click();
});

const log = (string) => {
	// eslint-disable-next-line no-console
	console.log(string);
};

const configureHotkeys = async () => {
	const settings = await browser.storage.sync.get(null);

	let modifierKey = 'alt';
	if (window.navigator.userAgentData) {
		// For newer browsers
		const { brands } = window.navigator.userAgentData;
		const isMac = brands.some((brand) =>
			brand.brand.toLowerCase().includes('mac'),
		);
		if (isMac) {
			modifierKey = 'option';
		}
	} else if (window.navigator.platform) {
		// For older browsers
		if (window.navigator.platform.indexOf('Mac') > -1) {
			modifierKey = 'option';
		}
	}

	hotkeys(settings.shortcutF1 || `ctrl+${modifierKey}+f1`, () => {
		alert('You pressed Ctrl+Mod+F1 shortcut');
	});

	hotkeys(settings.shortcutTilde || `~`, () => {
		alert('You pressed tilde');
	});
};

const init = async () => {
	const settings = await browser.storage.sync.get(null);
	if (settings.enabled) {
		if (settings.debugMode) {
			log(`[Verbatim] Enabled! Settings: ${JSON.stringify(settings)}`);
		}
		configureHotkeys();
	} else if (settings.debugMode) {
		log('[Verbatim] Disabled, doing nothing.');
	}
};

init();
