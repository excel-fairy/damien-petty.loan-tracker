/**
 * v0.1 - 20180706
 * Export a sheet in a spreadsheet on user's Google Drive
 *
 * Adapted from https://stackoverflow.com/questions/47289834/export-multiple-sheets-in-a-single-pdf
 * More info: https://stackoverflow.com/questions/11619805/using-the-google-drive-api-to-download-a-spreadsheet-in-csv-format
 *
 * @param {Object}  opts       (optional) Export options
 *                                Can contain any combination of fields
 *                                in following example:
 *                                {
 *                                    spreadSheetId: 'spreadSheetId',
 *                                    sheetId: 'sheetId',
 *                                    exportFolderId: 'folderId',
 *                                    exportFileName: 'file',
 *                                    range: {
 *                                        r1: 0,
 *                                        r2: 0,
 *                                        c1: 0,
 *                                        c2: 0
 *                                    }
 *                                    fileFormat: csv, pdf, etc.(?)
 *                                }
 */
function exportFile(opts) {

    opts = !!opts ? opts : {};

    // If a sheet ID was provided, open that sheet, otherwise assume script is
    // sheet-bound, and open the active spreadsheet.
    var ss = (opts.spreadSheetId) ? SpreadsheetApp.openById(opts.spreadSheetId) : SpreadsheetApp.getActiveSpreadsheet();

    // Get URL of spreadsheet, and remove the trailing 'edit'
    var url = ss.getUrl().replace(/edit$/,'');

    // Get folder containing spreadsheet, for later export
    // If folder ID is provided, use it. Otherwise export to
    // same folder as the one containing the spreadsheet
    var folder;
    if(opts.exportFolderId){
        folder = DriveApp.getFolderById(opts.exportFolderId);
    }
    else {
        var parents = DriveApp.getFileById(ss.getId()).getParents();
        if (parents.hasNext()) {
            folder = parents.next();
        }
        else {
            folder = DriveApp.getRootFolder();
        }
    }

    // Set range url parameters
    var rangeParameters = '';
    if(opts.range
        && opts.range.r1 && opts.range.r1 === parseInt(opts.range.r1, 10)
        && opts.range.r2 && opts.range.r2 === parseInt(opts.range.r2, 10)
        && opts.range.c1 && opts.range.c1 === parseInt(opts.range.c1, 10)
        && opts.range.c2 && opts.range.c2 === parseInt(opts.range.c2, 10))
        rangeParameters = '&r1=' + opts.range.r1 +
            '&r2=' + opts.range.r2 +
            '&c1=' + opts.range.c1 +
            '&c2=' + opts.range.c2;


    // If provided a sheetId, save it, otherwise save active sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    if(opts.sheetId){
        var sheets = ss.getSheets();
        for (var i=0; i<sheets.length; i++) {
            var currentSheet = sheets[i];
            if (opts.sheetId === currentSheet.getSheetId())
                sheet = currentSheet;
        }
    }

    // Additional parameters
    var fileFormat = 'pdf';
    if(opts.fileFormat && (opts.fileFormat === 'csv' || opts.fileFormat === 'pdf'))
        fileFormat = opts.fileFormat;
    var url_ext = 'export?exportFormat=' + fileFormat + '&format=' + fileFormat
        + '&gid=' + sheet.getSheetId()
        // following parameters are optional...
        + '&size=letter'      // paper size
        + '&portrait=true'    // orientation, false for landscape
        + '&fitw=true'        // fit to width, false for actual size
        + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
        + '&gridlines=false'  // hide gridlines
        + rangeParameters     // range
        + '&fzr=false';       // do not repeat row headers (frozen rows) on each page

    var options = {
        headers: {
            'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
        }
    };

    var response = UrlFetchApp.fetch(url + url_ext, options);

    var fileName;
    if(opts.exportFileName)
        fileName = opts.exportFileName + '.' + fileFormat;
    else
        fileName = ss.getName() + ' - ' + sheet.getName() + '.' + fileFormat;

    var blob = response.getBlob().setName(fileName);
    return folder.createFile(blob);
}
