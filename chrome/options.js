document.addEventListener('DOMContentLoaded', function () {

	chrome.storage.sync.get(null, function (settings) {
		if(settings.enabled == 0) {
			document.getElementById('cm_myonoffswitch').checked = false;
			console.log('Verbatim disabled');
		}
		else{
			document.getElementById('cm_myonoffswitch').checked = true;
			console.log('Verbatim enabled');
		}
		
		if(settings.copyselected == 1) {
			document.getElementById('copyselected').checked = true;
			console.log('Copy Selected enabled');
		}
		else{
			document.getElementById('copyselected').checked = false;
			console.log('Copy Selected disabled');
		}
	});
	
	document.querySelector('#cm_myonoffswitch').addEventListener('change', onOffHandler);
	document.querySelector('#copyselected').addEventListener('change', copySelectedHandler);
});

function onOffHandler(){
    if(document.getElementById('cm_myonoffswitch').checked){
		chrome.storage.sync.set({'enabled': 1}, function() {
			console.log('Cite Creator enabled');
		});
   }
   else{
      	chrome.storage.sync.set({'enabled': 0}, function() {
		  console.log('Cite Creator disabled');
		});
   }
}

function copySelectedHandler(){
    if(document.getElementById('copyselected').checked){
		chrome.storage.sync.set({'copyselected': 1}, function() {
			console.log('Copy Selected enabled');
		});
   }
   else{
      	chrome.storage.sync.set({'copyselected': 0}, function() {
		  console.log('Copy Selected disabled');
		});
   }
}
