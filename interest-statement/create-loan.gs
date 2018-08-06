
// var TEST_INTEREST_SHEET = {
//     sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test_Interest")
// };


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
    // LOANS_SHEET.sheet.appendRow(row); //TODO: position of appended row
}

// function appendTestInterests(data){
//     for (var i = 0; i < 12; i++){
//         var date = new Date();  //TODO: Current year ?
//         date.setDate(1);
//         date.setMonth(i);
//         var row = [];
//         var nbDaysInMonth = getNbDaysInMonth(date.getMonth(), date.getFullYear());
//         row[ColumnNames.letterToColumnStart0('A')] = date; //TODO: formatting
//         row[ColumnNames.letterToColumnStart0('B')] = nbDaysInMonth;
//         row[ColumnNames.letterToColumnStart0('C')] = nbDaysInMonth;
//         row[ColumnNames.letterToColumnStart0('D')] = 'TODO'; //TODO: Same as column A in Loans sheet
//         row[ColumnNames.letterToColumnStart0('E')] = data.interestRate;
//         row[ColumnNames.letterToColumnStart0('F')] = data.entityName;
//         row[ColumnNames.letterToColumnStart0('G')] = data.amountBorrowed;
//         row[ColumnNames.letterToColumnStart0('H')] = data.amountBorrowed * (data.interestRate / nbDaysInMonth); //TODO: Check helene's answer
//         row[ColumnNames.letterToColumnStart0('I')] = '';
//         row[ColumnNames.letterToColumnStart0('J')] = '';
//         row[ColumnNames.letterToColumnStart0('K')] = '';
//         row[ColumnNames.letterToColumnStart0('L')] = '';
//         row[ColumnNames.letterToColumnStart0('M')] = '';
//         row[ColumnNames.letterToColumnStart0('N')] = getMonthFullName(date.getMonth()+1);
//         row[ColumnNames.letterToColumnStart0('O')] = date.getFullYear();
//         row[ColumnNames.letterToColumnStart0('P')] = '';
//         row[ColumnNames.letterToColumnStart0('Q')] = '';
//         row[ColumnNames.letterToColumnStart0('R')] = '';
//         row[ColumnNames.letterToColumnStart0('S')] = '';
//         row[ColumnNames.letterToColumnStart0('T')] = '';
//         row[ColumnNames.letterToColumnStart0('U')] = '';
//         row[ColumnNames.letterToColumnStart0('V')] = '';
//         row[ColumnNames.letterToColumnStart0('W')] = '';
//         row[ColumnNames.letterToColumnStart0('X')] = '';
//         TEST_INTEREST_SHEET.sheet.appendRow(row);
//     }
// }

function getNbDaysInMonth (month, year) {
    return new Date(year, month, 0).getDate();
}

// Already exists
function getEntitiesNames(){
    var entities = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1+1).getValues();
    return entities.map(function(entity){return entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entityNameColumn];});
}

function getMonthFullName(month){
    var months = [
        'January', 'February', 'March', 'April', 'May', 'June', 'July',
        'August', 'September', 'October', 'November', 'December'
    ];
    return months[month-1];
}