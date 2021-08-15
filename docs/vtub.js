function getDriveDocs() {
    var docs = [];
    var folders = DriveApp.getFoldersByName('Tub');
    while (folders.hasNext()) {
        var folder = folders.next();
        var files = folder.getFiles();
        while (files.hasNext()) {
            var file = files.next();
            docs.push({ name: file.getName(), id: file.getId() });
        }
    }
    return docs;
}

function getDocContent(id) {
    var headings = [];
    var file = DocumentApp.openById(id);
    var body = file.getBody();
    var para = body.getParagraphs();
    for (var i = 0; i < para.length; i++) {
        var elem = para[i];
        if (elem.getHeading() === DocumentApp.ParagraphHeading.HEADING1
            || elem.getHeading() === DocumentApp.ParagraphHeading.HEADING2
            || elem.getHeading() === DocumentApp.ParagraphHeading.HEADING3
        ) {
            headings.push({ index: i, heading: elem.getHeading(), text: elem.getText() });
        }
    }
    return headings;
}

function getHeadingFromDoc(id, index) {
    var file = DocumentApp.openById(id);
    var body = file.getBody();
    var paragraphs = body.getParagraphs();
    var p = paragraphs[index].copy();

    var activeDoc = DocumentApp.getActiveDocument();
    var body = activeDoc.getBody();
    body.appendParagraph(p);
}
