/* ********************************
Verbatim by Aaron Hardy
v. 0.1 5/2/2021

ashtarcommunications@gmail.com
https://paperlessdebate.com
Provides keyboard shortcuts for Verbatim in Google Docs

Verbatim is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License 3.0 as published by
the Free Software Foundation.

Verbatim is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License 3 for more details.

For a copy of the GNU General Public License 3 see:
http://www.gnu.org/licenses/gpl-3.0.txt

*/

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
var editingIFrame = document.getElementsByClassName('docs-texteventtarget-iframe')[0];
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

window.addEventListener('message', (event) => {
    if (event.data === 'innerFrame') {
        console.log('received inner frame');
        innerFrame = event.source
    }
}, false);

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