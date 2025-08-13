chrome.extension.onMessage.addListener(
    async function (req, sender, sendResponse) {
        console.log('in the message');
        function paste() {
            document.execCommand('paste');
        }
        let [tab] = await chrome.tabs.query({ title: 'Speech - Google Docs' });
        chrome.tabs.executeScript(tab.id, { function: paste, allFrames: true });
        sendResponse('yep');
    }
);
