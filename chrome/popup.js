document.addEventListener('DOMContentLoaded', function () {

	chrome.storage.sync.get('enabled', function (settings) {
		if(settings.enabled == 0) {
			document.getElementById('cm_myonoffswitch').checked = false;
			console.log('Verbatim disabled');
		}
		else{
			document.getElementById('cm_myonoffswitch').checked = true;
			console.log('Verbatim enabled');
		}
	});

	document.querySelector('#cm_myonoffswitch').addEventListener('change', onOffHandler);

	document.getElementById('cite_popupoptionslink').href = chrome.extension.getURL("options.html");
	document.getElementById('cite_popupoptionslink').target = '_blank';
});

function onOffHandler(){
    if(document.getElementById('cm_myonoffswitch').checked){
		chrome.storage.sync.set({'enabled': 1}, function() {
			console.log('Verbatim enabled');
		});
   }
   else{
      	chrome.storage.sync.set({'enabled': 0}, function() {
		  console.log('Verbatim disabled');
		});
   }
}
