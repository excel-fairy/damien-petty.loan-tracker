/**
 * Main function
 * Called by custom menu
 */
function exportInterestStatements() {
    var entitiesNames = getEntitiesNames();
    var sheetUpdateInterval = 500; // Interval in ms between two entities switch. To let the spreadsheet to update itself. Not sure if needed
    var gSpreadSheetRateLimitingMinInterval = 6000; // Interval in ms between two exports. Google spreadsheet API (used to export sheet to PDF.
    // Returns HTTP 429 for rate limiting if too many requests are sent simultaneously
    for(var i=0; i < entitiesNames.length; i++){
        INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.entityCell).setValue(entitiesNames[i]);
        var totalValue = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.totalCell).getValue();
        Utilities.sleep(sheetUpdateInterval);
        if(totalValue !== 0){
            exportInterestStatementForCurrentEntity();
            Utilities.sleep(gSpreadSheetRateLimitingMinInterval-sheetUpdateInterval);
        }
    }
}

function exportInterestStatementForCurrentEntity(){
    var dateStr = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.dateCell).getValue();
    var entity = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.entityCell).getValue();
    var fileName = entity + ' - ' + INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.name + ' - ' + dateStr;
    var exportFolderId = getFolderToExportPdfTo(EXPORT_FOLDER_ID, dateStr).getId();

    var exportOptions = {
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.pdfExportRange
    };
    var exportedFile = ExportSpreadsheet.export(exportOptions);
    sendEmail(exportedFile);
}

function sendEmail(attachment) {
    var entityName = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.entityCell).getValue();
    var entity = getEntityFromName(entityName);
    if(!entity)
        SpreadsheetApp.getActiveSpreadsheet().toast('Entity ' + entityName + ' not found in entities list. No email sent');
    else {
        var recipient = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailAddressColumn];
        var subject = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailSubjectColumn];
        var message = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailBodyColumn];
        var carbonCopyEmailAddresses = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.carbonCopyEmailAddressesColumn];
        var emailOptions = {
            attachments: [attachment.getAs(MimeType.PDF)],
            name: 'Automatic loan tracker mail sender',
            cc: carbonCopyEmailAddresses
        };
        MailApp.sendEmail(recipient, subject, message, emailOptions);
    }
}

function getEntityFromName(entityName){
    var entities = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1+1).getValues();

    for (var i=0; i < entities.length; i++){
        if(entities[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entityNameColumn)] === entityName)
            return entities[i];
    }
    return null;
}