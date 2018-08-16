
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
        .setWidth(705)
        .setHeight(360);
    SpreadsheetApp.getUi().showDialog(htmlOutput);
}


function createLoanDebug(){
    insertLoanInLoansSheet({
        entityName:"Accordus Pty Ltd",
        amountBorrowed: 45,
        dateBorrowed: new Date(),
        dueDate: new Date(),
        interestRate: 12,
        borrowerEntity: "Antra group"
    })
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
    data.loanReference =  getIncrementedLoanReference(getLastLoanReferenceOfEntity(data.entityName));
    var rowToInsert = buildLoanToInsert(data);
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var lastEntityRow = getLastLoanOfEntityRow(data.entityName);
    var rangeRowToSet = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.lastLoansColumn)
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn));

    duplicateLastEntityRow(lastEntityRow);
    rangeRowToSet.setValues([rowToInsert]);
}

// Duplicate row to get all the data that won't be overwritten
function duplicateLastEntityRow(lastEntityRow){
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    loansOriginalSheet.insertRowAfter(lastEntityRow);
    var lastRangeRowOfEntity = loansOriginalSheet.getRange(lastEntityRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        loansOriginalSheet.getLastColumn()
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    var rangeRowToCopyDestination = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        loansOriginalSheet.getLastColumn()
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn) + 1);
    lastRangeRowOfEntity.copyTo(rangeRowToCopyDestination);
}

function buildLoanToInsert(data) {
    var row = [];
    row[ColumnNames.letterToColumnStart0('A')] = data.loanReference;
    row[ColumnNames.letterToColumnStart0('B')] = '';
    row[ColumnNames.letterToColumnStart0('C')] = data.entityName;
    row[ColumnNames.letterToColumnStart0('D')] = data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('E')] = data.dateBorrowed;
    row[ColumnNames.letterToColumnStart0('F')] = '';
    row[ColumnNames.letterToColumnStart0('G')] = data.dueDate;
    row[ColumnNames.letterToColumnStart0('H')] = data.interestRate + '%';
    row[ColumnNames.letterToColumnStart0('I')] = data.interestRate * data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('J')] = 'No';
    row[ColumnNames.letterToColumnStart0('K')] = '';
    row[ColumnNames.letterToColumnStart0('L')] = data.borrowerEntity;
    return row;
}


function getLastLoanReferenceOfEntity(entityName) {
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var lastLoanOfEntityRow = getLastLoanOfEntityRow(entityName);
    return loansOriginalSheet.getRange(lastLoanOfEntityRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn)).getValue();
}

function getLastLoanOfEntityRow(entityName) {
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var loansRange = loansOriginalSheet.getRange(2,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        loansOriginalSheet.getLastRow(),
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn) -
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn)+1);
    var allLoans = loansRange.getValues();
    var loanReference = getLastLoanOfEntity(entityName);
    for(var i=0; i < allLoans.length; i++){
        var currentLoanReference = allLoans[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn)];
        if( currentLoanReference === loanReference)
            return i + 1 + (INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow - 1);
    }
    return -1;
}

function getLastLoanOfEntity(entityName){
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var loansRange = loansOriginalSheet.getRange(2,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        loansOriginalSheet.getLastRow(),
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn) -
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn)+1);
    var allLoans = loansRange.getValues();
    var entityNameColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn);
    var loanReferenceColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn);
    allLoans = allLoans.filter(function (loan) {
        return loan[entityNameColS0] === entityName;
    });
    allLoans.sort(function (a, b) {
        return b[loanReferenceColS0].localeCompare([loanReferenceColS0]);
    });
    return allLoans[allLoans.length-1][loanReferenceColS0];
}

function getIncrementedLoanReference(loanReference) {
    // Split in two strings: letters, and digits (loan references
    // are a concatenation of a group of letters and a group of numbers
    var splittedLoanReference = loanReference.match(/[a-zA-Z]+|[0-9]+/g);
    var loanNumberStr = splittedLoanReference[1];
    var loanNumberStrLength = loanNumberStr.length;
    var loanNumber = parseInt(loanNumberStr, 10);
    var incrementedLoanNumber = loanNumber+1;
    var incrementedLoanNumberStr = ""+incrementedLoanNumber;
    while (incrementedLoanNumberStr.length < loanNumberStrLength) {
        incrementedLoanNumberStr = "0" + incrementedLoanNumberStr;
    }
    return splittedLoanReference[0] + incrementedLoanNumberStr;
}
