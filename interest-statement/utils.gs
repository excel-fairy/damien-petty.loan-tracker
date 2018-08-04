function getFolderToExportPdfTo(parentFolderId, date){
    var year = date.split(' ')[1];
    var month = date.split(' ')[0];
    return getChildFolderByNameAndCreateIfNotExist(getChildFolderByNameAndCreateIfNotExist(parentFolderId, year).getId(), month);
}

function getChildFolderByNameAndCreateIfNotExist(parentFolderId, childFolderName){
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var folders = parentFolder.getFoldersByName(childFolderName);
    if (folders.hasNext()){ // Return first child folder with specified name
        return folders.next();
    }
    return parentFolder.createFolder(childFolderName); // Create child folder with specified name and return it
}
