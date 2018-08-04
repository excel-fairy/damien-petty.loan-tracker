function getMonthFolderNameFromMonthName(month){
    var months = {
        January: 1,
        February: 2,
        March: 3,
        April: 4,
        May: 5,
        June: 6,
        July: 7,
        August: 8,
        September: 9,
        October: 10,
        November: 11,
        December: 12
    };
    return months[month] + "." + month;
}

function getFolderToExportPdfTo(parentFolderId, date){
    var year = date.split(' ')[1];
    var month = getMonthFolderNameFromMonthName(date.split(' ')[0]);
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