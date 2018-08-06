
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
    SpreadsheetApp.getUi().alert ("Loan is being imported. Please wait for it to be fully created");
    appendLoanToLoansSheet(data);
    appendTestInterests(data);
}

function appendLoanToLoansSheet(data){
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
    SpreadsheetApp.openById(LOAN_TRACKER_SPREADSHEET_ID).getSheetByName(LOAN_TRACKER_SPREADSHEET.loansSheet.name).appendRow(row);
}
