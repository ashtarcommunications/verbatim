$(function () {
	console.log('sending from inner frame');
	window.parent.parent.parent.postMessage('innerFrame', '*');
	window.addEventListener(
		'message',
		(event) => {
			if (event.data === 'sendToSpeech') {
				console.log('sending to speech');
				sendToSpeech();
			}
			if (event.data === 'extensioninstalled') {
				console.log('extension installed');
			}
		},
		false,
	);

	$('#new-speech').click(newSpeech);
	$('#choose-speech').click(chooseSpeech);
	$('#send').click(sendToSpeech);
	$('#condense').click(condense);
	$('#copy').click(copySelected);
	$('#paste').click(pasteSelected);
	$('#select').click(selectHeading);
	$('#moveup').click(moveUp);
	$('#movedown').click(moveDown);
	$('#styles').click(styles);

	$('#pocket').click(setHeading);
	$('#hat').click(setHeading);
	$('#block').click(setHeading);
	$('#tag').click(setHeading);
	$('#cite').click(setFormatting);
	$('#underline').click(setFormatting);
	$('#emphasis').click(setFormatting);
	$('#highlight').click(setFormatting);
	$('#clear').click(setFormatting);

	$('#shrink').click(shrink);
	$('#invisibility-on').click(invisibilityOn);
	$('#invisibility-off').click(invisibilityOff);

	$('#highlight-color').change(changeHighlight);

	$('#search-button').click(search);

	$('#get-files').click(getFiles);
	$('#get-doc').click(getDoc);
	$('#get-heading').click(getHeadingFromDoc);
	$('#create-vtub').click(createVtub);
	// $('#export').click(exportAsDocx);

	$('.tabs li').click(switchTab);
});

const run = (
	successHandler,
	failureHandler = (err) => $('#error-message').text(err),
	userObject = this,
) => {
	return google.script.run
		.withSuccessHandler(successHandler)
		.withFailureHandler(failureHandler)
		.withUserObject(userObject);
};

let settings;
const settingsSuccess = (s) => {
	console.log(s);
	settings = s;
	const activeSpeech =
		settings && settings.ACTIVE_SPEECH
			? JSON.parse(settings.ACTIVE_SPEECH)
			: null;
	if (activeSpeech) {
		$('#current-speech').html(
			`<p>Current Speech:</p><p><a target="_blank" href="${activeSpeech.url}">${activeSpeech.name}</a></p>`,
		);
	}
};
// run(settingsSuccess).getProperties();

const sendToSpeech = () => {
	this.disabled = true;
	$('#error').remove();
	run().sendToSpeech();
	this.disabled = false;
};

Mousetrap.bind('`', function (e) {
	alert('You pressed tilde');
	return false;
});

const newSpeech = () => {
	console.log('running new speech');
	const successHandler = (documentId) => {
		console.log(documentId);
		// const link = document.createElement('a');
		// link.href = `https://docs.google.com/document/d/${documentId}`;
		// link.target = '_blank';
	};
	run(sucessHandler).newSpeech();
};
const chooseSpeech = () => run().showPicker();
const setHeading = (e) => run().setHeading(e.currentTarget.id);
const setFormatting = (e) => run().setFormatting(e.currentTarget.id);
const styles = () => run().restoreStyles();
const shrink = () => run().shrink();
const invisibilityOn = () => run().invisibilityOn();
const invisibilityOff = () => run().invisibilityOff();
const condense = () => run().condense();
const selectHeading = () => run().selectHeading();
const moveUp = () => run().moveUp();
const moveDown = () => run().moveDown();
const copySelected = () => run().copySelected();
const pasteSelected = () => run().pasteSelected();
const changeHighlight = (e) => run().changeHighlight(e.currentTarget.value);
const search = () => run().searchDrive($('#search-query').val());
const getFiles = () => {
	const successHandler = (files) => files.forEach((f) => console.log(f));
	run(successHandler).getDriveDocs();
};
const getDoc = () => {
	const successHandler = (headings) => headings.forEach((h) => console.log(h));
	run(successHandler).getDocContent(
		'1Dj1loasY1jM38Hl3dOujHZs9Y0gKM93iUoTFPVtbxBc',
	);
};
const getHeadingFromDoc = () => {
	run().getHeadingFromDoc('1Dj1loasY1jM38Hl3dOujHZs9Y0gKM93iUoTFPVtbxBc', 1);
};
const createVtub = () => run().createVtub();

const switchTab = function() {
	const targetTab = $(this).data('tab');

	// Remove active class from all tabs and pages
	$('.tabs li').removeClass('active');
	$('.page').removeClass('active');

	// Add active class to clicked tab
	$(this).addClass('active');

	// Show corresponding page
	$('#' + targetTab).addClass('active');
};
// const exportAsDocx = () => {
// 	const successHandler = async (base64Data) => {
// 		// Convert base64 to array buffer
// 		const byteChars = atob(base64Data);
// 		const byteNumbers = new Array(byteChars.length);
// 		for (let i = 0; i < byteChars.length; i++) {
// 			byteNumbers[i] = byteChars.charCodeAt(i);
// 		}
// 		const byteArray = new Uint8Array(byteNumbers);

// 		// Load zip
// 		const zip = await JSZip.loadAsync(byteArray);

// 		// =====================
// 		// 1. settings.xml.rels
// 		// =====================
// 		const relsPath = 'word/_rels/settings.xml.rels';
// 		const targetPath =
// 			'file:///C:/Users/ashtar/AppData/Roaming/Microsoft/Templates/Debate.dotm';
// 		let relId = 'rId1'; // default id if creating new

// 		let relsDoc;
// 		if (zip.file(relsPath)) {
// 			let relsXml = await zip.file(relsPath).async('string');
// 			relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
// 		} else {
// 			relsDoc = new DOMParser().parseFromString(
// 				'<?xml version="1.0" encoding="UTF-8"?>' +
// 					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
// 				'application/xml',
// 			);
// 		}

// 		// Find or create attachedTemplate relationship
// 		let rels = relsDoc.getElementsByTagName('Relationship');
// 		let templateRel = null;
// 		for (let i = 0; i < rels.length; i++) {
// 			if (
// 				rels[i].getAttribute('Type') ===
// 				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate'
// 			) {
// 				templateRel = rels[i];
// 				relId = rels[i].getAttribute('Id'); // reuse its Id
// 				break;
// 			}
// 		}

// 		if (templateRel) {
// 			templateRel.setAttribute('Target', targetPath);
// 		} else {
// 			const newRel = relsDoc.createElement('Relationship');
// 			newRel.setAttribute('Id', relId);
// 			newRel.setAttribute(
// 				'Type',
// 				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate',
// 			);
// 			newRel.setAttribute('Target', targetPath);
// 			relsDoc.documentElement.appendChild(newRel);
// 		}

// 		const updatedRelsXml = new XMLSerializer().serializeToString(relsDoc);
// 		zip.file(relsPath, updatedRelsXml);

// 		// =====================
// 		// 2. word/settings.xml
// 		// =====================
// 		const settingsPath = 'word/settings.xml';
// 		let settingsDoc;

// 		if (zip.file(settingsPath)) {
// 			let settingsXml = await zip.file(settingsPath).async('string');
// 			settingsDoc = new DOMParser().parseFromString(
// 				settingsXml,
// 				'application/xml',
// 			);
// 		} else {
// 			// If missing entirely (rare from Google Docs), create a basic skeleton
// 			settingsDoc = new DOMParser().parseFromString(
// 				'<?xml version="1.0" encoding="UTF-8"?>' +
// 					'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
// 					'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>',
// 				'application/xml',
// 			);
// 		}

// 		const wNS =
// 			'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
// 		const rNS =
// 			'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// 		// Look for existing w:attachedTemplate
// 		let attachedTemplate = settingsDoc.getElementsByTagNameNS(
// 			wNS,
// 			'attachedTemplate',
// 		)[0];

// 		if (attachedTemplate) {
// 			// Update the r:id to match relationship
// 			attachedTemplate.setAttributeNS(rNS, 'r:id', relId);
// 		} else {
// 			// Create new element
// 			attachedTemplate = settingsDoc.createElementNS(
// 				wNS,
// 				'w:attachedTemplate',
// 			);
// 			attachedTemplate.setAttributeNS(rNS, 'r:id', relId);

// 			// Append it (placing it at the root is fine, usually right under <w:settings>)
// 			settingsDoc.documentElement.appendChild(attachedTemplate);
// 		}

// 		const updatedSettingsXml = new XMLSerializer().serializeToString(
// 			settingsDoc,
// 		);
// 		zip.file(settingsPath, updatedSettingsXml);

// 		// Re-zip and trigger download
// 		const blob = await zip.generateAsync({ type: 'blob' });
// 		const url = URL.createObjectURL(blob);
// 		const a = document.createElement('a');
// 		a.href = url;
// 		a.download = 'export.docx';
// 		a.click();
// 		URL.revokeObjectURL(url);
// 	};
// 	run(successHandler).exportAsDocx();
// };

// const detectExtension = () => {
//     setTimeout(() => {
//         const div = document.getElementById('extensionintalled');
//         if (!div) { alert('no ext'); }
//     }, 5000);
// }
// detectExtension();
