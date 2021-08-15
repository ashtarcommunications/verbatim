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
// heading3Style[DocumentApp.Attribute.UNDERLINE] = true;
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

function condense() {
    // Credit to Mark Crimmins' Remove Line Breaks add-on
    // https://workspace.google.com/marketplace/app/remove_line_breaks/253339336957
    var selection = DocumentApp.getActiveDocument().getSelection();

    if (!selection) {
        DocumentApp.getUi().alert('No text is selected.  Please select a text region in which line breaks are to be removed.');
        return;
    }

    /**
     * selectedElements is a list of Body components (such as Paragraphs, List Items, Tables, etc., 
     * but the first and last items might be partially selected, in which case they are Ranges
     * whose Parents are the components.
     */
    var selectedElements = selection.getSelectedElements();

    /**
     * We will merge a paragraph with a preceding one only if both are of the NORMAL
     * Heading type (Normal text rather than Title, Heading 1, etc.).  We consider list items
     * "normal" too.  We start with false (non-normal) so that we don't try to merge the 
     * first paragraph with a nonexistent predecessor.
     */
    var normalPredecessor = false;
    var prevType = DocumentApp.ElementType.PARAGRAPH;

    /**
     * Main loop.  Get next element in the selection.  If it's a normal-text paragraph or a list-item, then 
     * after removing CR and LF characters we will merge it with its predecessor if the predecessor is normal 
     * and of the same type.
     */
    var len = selectedElements.length;
    for (var i = 0; i < len; i++) {
        var nextElementOrRange = selectedElements[i];

        /* If the next element is only partially selected, get the whole element */
        if (nextElementOrRange.isPartial()) {
            var nextElement = nextElementOrRange.getElement().getParent();
        } else {
            var nextElement = nextElementOrRange.getElement();
        }

        /* What is the type of the next element? */
        var nextType = nextElement.getType();
        if ((nextType == DocumentApp.ElementType.PARAGRAPH && nextElement.getHeading() == DocumentApp.ParagraphHeading.NORMAL) || nextType == DocumentApp.ElementType.LIST_ITEM) {

            /* We have a normal paragraph or a list item now */

            var nextText = nextElement.editAsText();

            /* Trim any preceding and trailing spaces. */
            nextText.replaceText('^\\s*', '');
            nextText.replaceText('\\s*$', '');

            if (nextText.getText() != "") {
                /* A nonempty normal paragraph or list item. */

                /* Replace carriage returns and newlines (characters) with spaces
                   Unfortunately, replaceText doesn't actually work, as at least CR 
                   seems not to be recognized as \r, despite having ascii code 13 */

                /* nextText.replaceText('[\\r\\n]',' '); */

                /* So as a kludge, we seek for CR and LF characters (ascii 10 and 13) 
                   and delete them, inserting spaces in their places. */
                var str = nextText.getText();
                var str2 = '';
                if (nextElementOrRange.isPartial()) {
                    endindex = nextElementOrRange.getEndOffsetInclusive();
                    startindex = nextElementOrRange.getStartOffset();
                } else {
                    endindex = str.length - 1;
                    startindex = 0;
                }
                for (var j = endindex; j >= startindex; j--) {
                    var ascii = str.charCodeAt(j);
                    if (ascii == 10 || ascii == 13) {
                        nextText.deleteText(j, j);
                        nextText.insertText(j, ' ');
                    }
                }

                /* If predecessor is normal and of same type, prepend a space and merge. */

                if (normalPredecessor && nextType == prevType) {
                    nextElement.asText().insertText(0, ' ');
                    nextElement.merge();
                }
                normalPredecessor = true;
                prevType = nextType;
            } else {
                /* Blank paragraph or list item. Remove it if it's not the end of the document
                 * and preserve values of normalPredecessor and prevType. 
                 */
                if (i < len - 1) {
                    /* You can't remove the last element in a document. */
                    nextElement.removeFromParent();
                }
            }
        } else {
            /* nextElement is not a normal paragraph or list item. */
            normalPredecessor = false;
        }
    }
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
                var color = getSetting('HIGHLIGHT_COLOR') || '#ffff00';
                partial ? e.setBackgroundColor(start, end, color) : e.setBackgroundColor(color);
                // partial ? e.setForegroundColor(start, end, '#000000') : e.setForegroundColor('#000000');
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

function changeHighlight(color) {
    Logger.log(color);
    var prop = PropertiesService.getUserProperties();
    prop.setProperty('HIGHLIGHT_COLOR', color);
}

function shrink() {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var element = cursor.getElement();
    var p = element.getParent().asText();
    var length = p.getText().length;
    for (var i = 0; i < length; i++) {
        if (!p.isUnderline(i)) {
            p.setFontSize(i, i, 8);
        }
    }
}

function invisibilityOn() {
    // TODO - rewrite to go paragraph by paragrah and skip headings
    // TODO - add a safety check to bail if document is too long
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();
    var length = text.getText().length;
    for (var i = 0; i < length; i++) {
        if (!text.getBackgroundColor(i)) {
            text.setForegroundColor(i, i, '#ffffff');
        }
    }
}

function invisibilityOff() {
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();
    var length = text.getText().length;
    for (var i = 0; i < length; i++) {
        if (!text.getBackgroundColor(i)) {
            text.setForegroundColor(i, i, '#000000');
        }
    }
}
