/**
 * Main function
 * Called by custom menu
 */
function exportInterestStatements() {
    var entitiesNames = getEntitiesNames();
    var sheetUpdateInterval = 500; // Interval in ms between two entities switch. To let the spreadsheet to update itself. Not sure if needed
    var gSpreadSheetRateLimitingMinInterval = 6000; // Interval in ms between two exports. Google spreadsheet API (used to export sheet to PDF.
                                                    // Returns HTTP 429 for rate limiting if too many requests are sent simultaneously
    var currentlyExportingEntity = getCurrentlyExportingEntity();
    var startIndex = 0;
    if(currentlyExportingEntity !== '')
        startIndex = entitiesNames.indexOf(currentlyExportingEntity);
    var exportExecutionStartDate = new Date();
    for(var i = startIndex; i < entitiesNames.length; i++){
        var entityName = entitiesNames[i];
        INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.entityCell).setValue(entityName);
        setCurrentlyExportingEntity(entityName);

        // Stops if script execution is becoming too close from the GAS limit per script (limit is 5min, stops at 4m30s)
        if(isTimeUp(exportExecutionStartDate)){
            var currentlyExportingEntityIndex = entitiesNames.indexOf(getCurrentlyExportingEntity());
            var lastExportedEntity = currentlyExportingEntityIndex > 0 ? currentlyExportingEntityIndex - 1 : 0;
            updateExportStatus(false);
            SpreadsheetApp.getActiveSpreadsheet().toast('Script execution is too long and had to stop. The last ' +
                'exported entity is ' + lastExportedEntity + '. Next execution will start from here');
            return;
        }
        else
            updateExportStatus(true);

        var totalValue = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.totalCell).getValue();
        Utilities.sleep(sheetUpdateInterval);
        if(totalValue !== 0){
            exportInterestStatementForCurrentEntity();
            Utilities.sleep(gSpreadSheetRateLimitingMinInterval-sheetUpdateInterval);
        }
        updateExportStatus(false)
    }
    setCurrentlyExportingEntity('');
    updateExportStatus(false)
}

function getCurrentlyExportingEntity() {
    return INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingEntityCell).getValue();
}

function setCurrentlyExportingEntity(entityName) {
    return INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingEntityCell).setValue(entityName);
}

/**
 * Update the GUI display of the export progress
 * @param executionOnGoing Is the export ongoing (true) or has it stopped (false) ?
 */
function updateExportStatus(executionOnGoing) {
    var textToWrite;
    if(executionOnGoing)
        textToWrite = "Script in progress";
    else {
        textToWrite = 'All Entities exported!';
        var lastExportedEntity = getCurrentlyExportingEntity();
        if(lastExportedEntity !== '')
            textToWrite = 'Exported until entity ' + lastExportedEntity + ', hit "Export & Send" again to proceed with the export';
    }
    INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.exportStatusCell).setValue(textToWrite);
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