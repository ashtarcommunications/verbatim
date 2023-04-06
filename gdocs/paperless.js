function sendToSpeech() {
    var activeDoc = DocumentApp.getActiveDocument();
    selectHeading();
    var selection = activeDoc.getSelection();

    let activeSpeech = getProperty('ACTIVE_SPEECH');
    if (activeSpeech) {
        activeSpeech = JSON.parse(activeSpeech);
    } else {
        return false;
    }
    var speechDoc = DocumentApp.openById(activeSpeech.id);
    var speechBody = speechDoc.getBody();

    var elements = selection.getRangeElements();

    for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        var e = element.getElement();
        var p = e.asParagraph();
        speechBody.appendParagraph(p.copy());
    }

    var paragraph = speechDoc.getBody().appendParagraph('');
    // var position = speechDoc.newPosition(paragraph, 0);
    // speechDoc.setCursor(position);
}

function copySelected() {
    var selection = DocumentApp.getActiveDocument().getSelection();
    Logger.log(selection);
    var str = JSON.stringify(selection);
    Logger.log(str);
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('SELECTED', str);
}

function pasteSelected() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var str = scriptProperties.getProperty('SELECTED');
    var selection = JSON.parse(str);
    DocumentApp.getUi().alert(JSON.stringify(selection));

    var body = DocumentApp.getActiveDocument().getBody();
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        var e = element.getElement();
        var p = e.asParagraph();
        body.appendParagraph(p.copy());
    }
}

function selectHeading() {
    const headings = {
        'TITLE': 1,
        'SUBTITLE': 1,
        'HEADING1': 1,
        'HEADING2': 2,
        'HEADING3': 3,
        'HEADING4': 4,
        'HEADING5': 5,
        'HEADING6': 6,
        'NORMAL': 7,
    };

    var doc = DocumentApp.getActiveDocument();
    var range = doc.newRange();
    var cursor = doc.getCursor();
    var currentElement = cursor.getElement();
    if (currentElement.getType() === DocumentApp.ElementType.TEXT) {
        currentElement = currentElement.getParent();
    }
    range.addElement(currentElement);

    var heading = currentElement.getHeading();
    var startHeading = heading;

    if (headings[heading] > 6) {
        var startElement = currentElement;
        while (startElement) {
            const prevSibling = startElement.getPreviousSibling();
            if (prevSibling) {
                range.addElement(prevSibling);
                const prevHeading = prevSibling.getHeading();
                startHeading = prevHeading;
                if (headings[prevSibling.getHeading()] < 7) {
                    startElement = null;
                    break;
                } else {
                    startElement = prevSibling;
                }
            } else {
                startElement = null;
                break;
            }
        }
    }

    var endElement = currentElement;
    while (endElement) {
        const nextSibling = endElement.getNextSibling();
        if (nextSibling) {
            if (headings[nextSibling.getHeading()] <= headings[startHeading]) {
                endElement = null;
                break
            } else {
                range.addElement(nextSibling);
                endElement = nextSibling;
            }
        } else {
            endElement = null;
            break
        }
    }

    doc.setSelection(range.build());
}

function moveUp() {
    selectHeading();
    var body = DocumentApp.getActiveDocument().getBody();
    var selection = DocumentApp.getActiveDocument().getSelection();
    var elements = selection.getRangeElements();
    var firstElement = elements[0].getElement();
    var firstHeading = firstElement.getHeading();
    var firstIndex = body.getChildIndex(firstElement);

    var elementToInsertBefore;

    for (var i = firstIndex - 1; i = 0; i--) {
        var e = body.getChild(i);
        if (e.getType() === DocumentApp.ElementType.PARAGRAPH) {
            if (e.getHeading() === firstHeading) {
                elementToInsertBefore = e;
                break;
            }
        }
    }

    elements.forEach(element => {
        var e = element.getElement();
        e.removeFromParent();
        body.insertParagraph(body.getChildIndex(elementToInsertBefore) - 1, e);
    });
}

const searchDrive = (q) => {
    // TODO - escape quotes in q
    var files = DriveApp.searchFiles(`title contains "${q}"`);
    // var fulltext = DriveApp.searchFiles(`fulltext contains "${q}"`);
    while (files.hasNext()) {
        let file = files.next();
        Logger.log(file.getName());
    }
    // while (fulltext.hasNext()) {
    //     let file = fulltext.next();
    //     Logger.log(file.getName());
    // }
}

const newSpeech = () => {
    const speechName = 'Speech 2AC 8-9 6PM';
    var doc = DocumentApp.create(speechName);
    const id = doc.getId();
    Logger.log(id);
    var html = "<script>window.open(`https://docs.doogle.com/document/d/${id}`);google.script.host.close();</script>";
    var html = "<script>var link=document.createElement('a');link.href=`https://docs.doogle.com/document/d/${id}`;link.target='_blank';link.click();google.script.host.close();</script>";
    var ui = HtmlService.createHtmlOutput(html);
    DocumentApp.getUi().showModalDialog(ui, 'Open Doc');
    return id;
}
