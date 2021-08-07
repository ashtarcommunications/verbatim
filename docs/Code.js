function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Start', 'showSidebar')
        .addToUi();

    showSidebar();
    setStyles();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('Verbatim');
    DocumentApp.getUi().showSidebar(ui);
}

var heading1Style = {};
heading1Style[DocumentApp.Attribute.BOLD] = true;
heading1Style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
heading1Style[DocumentApp.Attribute.FONT_SIZE] = '32';
heading1Style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
heading1Style[DocumentApp.Attribute.BORDER_WIDTH] = '1';
heading1Style[DocumentApp.Attribute.BORDER_COLOR] = 'black';

var heading2Style = {};
heading2Style[DocumentApp.Attribute.BOLD] = true;
heading2Style[DocumentApp.Attribute.UNDERLINE] = true;
heading2Style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
heading2Style[DocumentApp.Attribute.FONT_SIZE] = '24';
heading2Style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

var heading3Style = {};
heading3Style[DocumentApp.Attribute.BOLD] = true;
heading3Style[DocumentApp.Attribute.UNDERLINE] = true;
heading3Style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
heading3Style[DocumentApp.Attribute.FONT_SIZE] = '18';
heading3Style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

var heading4Style = {};
heading4Style[DocumentApp.Attribute.BOLD] = true;
heading4Style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
heading4Style[DocumentApp.Attribute.FONT_SIZE] = '12';

function setStyles() {
    var body = DocumentApp.getActiveDocument().getBody();

    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, heading1Style);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, heading2Style);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, heading3Style);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING4, heading4Style);
}

function restoreStyles() {
    var body = DocumentApp.getActiveDocument().getBody();
    var para = body.getParagraphs();
    for (var i in para) {
        switch (para[i].getHeading()) {
            case DocumentApp.ParagraphHeading.HEADING1:
                para[i].setAttributes(heading1Style);
                break;
            case DocumentApp.ParagraphHeading.HEADING2:
                para[i].setAttributes(heading2Style);
                break;
            case DocumentApp.ParagraphHeading.HEADING3:
                para[i].setAttributes(heading3Style);
                break;
            case DocumentApp.ParagraphHeading.HEADING4:
                para[i].setAttributes(heading4Style);
                break;
        }
    }
    var body = DocumentApp.getActiveDocument().getBody();

    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, heading1Style);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, heading2Style);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, heading3Style);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING4, heading4Style);
}

function sendToSpeech() {
    var activeDoc = DocumentApp.getActiveDocument();
    selectHeading();
    var selection = activeDoc.getSelection();

    var speechId = '1hmhWoqEIRYMEhafTnGmPmzoJ_BDbqzJbj1aCKrhQ2yo';
    var speechDoc = DocumentApp.openById(speechId);
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

function condense() {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (selection) {
        var elements = selection.getRangeElements();
        for (var i = 0; i < elements.length; i++) {
            var element = elements[i];

            // Only deal with text elements
            if (element.getElement().editAsText) {
                // var text = element.getElement().editAsText();

                // if (element.isPartial()) {
                //     text.replaceText("\\n", " ");
                //     text.replaceText("\\p{Cc}+", " ")
                //     text.replaceText("\\v+", " ");
                // } else {
                //     // Deal with fully selected text
                //     text.replaceText("\\n", " ");
                //     text.replaceText("\\p{Cc}+", " ")
                //     text.replaceText("\\v+", " ");
                // }
              var text2 = element.getElement().getText();
              text2 = text2.replace('\n', ' ');
              element.getElement().setText(text2);
            }
        }
    } else {
        DocumentApp.getUi().alert('Select something first!');
    }
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

    // var startElement = firstElement.copy();
    // while (startElement) {
    //     const prevSibling = startElement.getPreviousSibling();
    //     if (prevSibling) {
    //         const prevHeading = prevSibling.getHeading();
    //         if (prevHeading === firstHeading) {
    //             elementToInsertBefore = prevSibling;
    //             startElement = null;
    //             break;
    //         } else {
    //             startElement = prevSibling;
    //         }
    //     } else {
    //         startElement = null;
    //         break;
    //     }
    // }

    // var indexToInsertBefore = body.getChildIndex(elementToInsertBefore);

    elements.forEach(element => {
        var e = element.getElement();
        e.removeFromParent();
        body.insertParagraph(body.getChildIndex(elementToInsertBefore) - 1, e);
    });
}

function setHeading(heading) {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var element = cursor.getElement();
    var p = element.getParent().asParagraph();
    Logger.log(p.getAttributes());
    switch (heading) {
        case 'pocket':
            p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
            break;
        case 'hat':
            p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
            break;
        case 'block':
            p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
            break;
        case 'tag':
            p.setHeading(DocumentApp.ParagraphHeading.HEADING4);
            break;
        case 'normal':
        default:
            p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }
}

function setFormatting(format) {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (!selection) { return false; }
    var elements = selection.getRangeElements();
    elements.forEach(element => {
        var partial = element.isPartial();
        var start = element.getStartOffset();
        var end = element.getEndOffsetInclusive();
        var e = element.getElement().asText();

        switch (format) {
            case 'cite':
                partial ? e.setBold(start, end, true) : e.setBold(true);
                partial ? e.setFontSize(start, end, 13) : e.setFontSize(13);
                break;
            case 'underline':
                partial ? e.setUnderline(start, end, true) : e.setUnderline(true);
                break;
            case 'emphasis':
                partial ? e.setBold(start, end, true) : e.setBold(true);
                break;
            case 'highlight':
                partial ? e.setBackgroundColor(start, end, '#ffff00') : e.setBackgroundColor('#ffff00');
                partial ? e.setForegroundColor(start, end, '#000000') : e.setForegroundColor('#000000');
                break;
            case 'clear':
            default:
                e.setBold(false);
                e.setUnderline(false);
                e.setFontSize(11);
                e.setBackgroundColor(null);
                e.setForegroundColor(null);
        }
    });
}
