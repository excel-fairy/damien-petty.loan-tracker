
/**
 * Called by custom menu
 */
function openCreateLoanPopup() {
    var htmlTemplate = HtmlService.createTemplateFromFile('createloan');
    htmlTemplate.data = {
        entities: getEntitiesNames(),
        borrowers: ['Antra Group', 'Ray Petty']
    };
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Import loan')
        .setWidth(900)
        .setHeight(500);
    SpreadsheetApp.getUi().showDialog(htmlOutput);
}


/**
 * Main function
 * Called by HTML button in popup
 */
function createLoan(data) {
    SpreadsheetApp.getUi().alert ('Loan is being imported. It will appear in the "Loans" tab shortly');
    insertLoanInLoansSheet(data);
}

function insertLoanInLoansSheet(data){
    data.loanReference =  getLastLoanReferenceOfEntity(data.entityName);
    var rowToInsert = buildLoanToInsert(data);
    var loansOriginalSheet = getLoansOriginalSheet();
    loansOriginalSheet.insertRowAfter(lastEntityRow);
    var rangeRowToSet = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.lastLoansColumn)
        - ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    rangeRowToSet.setValues([rowToInsert]);
}

function duplicateLastEntoityRow(lastEntityRow,){
    var loansOriginalSheet = getLoansOriginalSheet();
    var lastRangeRowOfEntity = loansOriginalSheet.getRange(lastEntityRow,
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        loansOriginalSheet.getLastColumn()
        - ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    var rangeRowToCopyDestination = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        ColumnNames.letterToColumn(loansOriginalSheet.lastLoansColumn)
        - ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    lastRangeRowOfEntity.copyTo(rangeRowToCopyDestination);
}

function buildLoanToInsert(data) {
    var row = [];
    row[ColumnNames.letterToColumnStart0('A')] = 'TODO'; //TODO
    row[ColumnNames.letterToColumnStart0('B')] = '';
    row[ColumnNames.letterToColumnStart0('C')] = data.entityName;
    row[ColumnNames.letterToColumnStart0('D')] = data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('E')] = data.dateBorrowed;
    row[ColumnNames.letterToColumnStart0('F')] = '☐';
    row[ColumnNames.letterToColumnStart0('G')] = data.dueDate;
    row[ColumnNames.letterToColumnStart0('H')] = data.interestRate;
    row[ColumnNames.letterToColumnStart0('I')] = data.interestRate * data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('J')] = 'No';
    row[ColumnNames.letterToColumnStart0('K')] = '';
    row[ColumnNames.letterToColumnStart0('L')] = data.borrowerEntity;
    return row;
}

function getLastLoanReferenceOfEntity(entityName) {
    var loansOriginalSheet = getLoansOriginalSheet();
    var lastLoanOfEntityRow = getLastLoanOfEntityRow(entityName);
    return loansOriginalSheet.getRange(lastLoanOfEntityRow, LOAN_TRACKER_SPREADSHEET.loansSheet.loanReferenceColumn).getValue();
}

function getLastLoanOfEntityRow(entityName){
    var loansOriginalSheet = getLoansOriginalSheet();
    var loans = loansOriginalSheet.getRange(2,
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn),
        loansOriginalSheet.getLastRow(),
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.entityNameColumn) -
        ColumnNames.letterToColumn(LOAN_TRACKER_SPREADSHEET.loansSheet.firstLoansColumn)+1);
    //TODO

}

function getLoansOriginalSheet() {
    return SpreadsheetApp.openById(LOAN_TRACKER_SPREADSHEET_ID).getSheetByName(LOAN_TRACKER_SPREADSHEET.loansSheet.name);
}


