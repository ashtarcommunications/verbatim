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

function getFilesInFolder(folder) {
    var vtub = {};
    var files = folder.getFiles();
    while (files.hasNext()) {
        var file = files.next();
        var fileId = file.getId();
        vtub[fileId] = { name: file.getName(), type: 'file', headings: getDocContent(fileId) };
    }
    return vtub;
}

function createVtub() {
    var vtub = {};
    var folder = DriveApp.getFolderById('1NZ4Rgx6-nU9aYjPCCA7A-ZqYgLdcGSFq');
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
        var subfolder = subfolders.next();
        var subfolderId = subfolder.getId();
        vtub[subfolderId] = { name: subfolder.getName(), type: 'folder', ...getFilesInFolder(subfolder) };
        var subsubfolders = subfolder.getFolders();
        while (subsubfolders.hasNext()) {
            var subsubfolder = subsubfolders.next();
            var subsubfolderId = subsubfolder.getId();
            vtub[subfolderId][subsubfolderId] = { name: subsubfolder.getName(), type: 'folder', ...getFilesInFolder(subsubfolder) };
            var subsubsubfolders = subsubfolder.getFolders();
            while (subsubsubfolders.hasNext()) {
                var subsubsubfolder = subsubsubfolders.next();
                var subsubsubfolderId = subsubsubfolder.getId();
                vtub[subfolderId][subsubfolderId][subsubsubfolderId] = { name: subsubsubfolder.getName(), type: 'folder', ...getFilesInFolder(subsubsubfolder) };
            }
        }
    }
    vtub = { ...vtub, ...getFilesInFolder(folder) };
    DriveApp.createFile('vtub.json', JSON.stringify(vtub));
}
