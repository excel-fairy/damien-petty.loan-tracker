
/**
 * Called by custom menu
 */
function openCreateLoanPopup() {
    var htmlTemplate = HtmlService.createTemplateFromFile('createloan');
    htmlTemplate.data = {
        entities: getEntitiesNames(),
        borrowers: ['Antra Group', 'Ray Petty', 'Fundsquire Pty Ltd']
    };
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Import loan')
        .setWidth(705)
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
    // Override loanReference (autocomputed) only if the entity is none of the below
    if(data.loanReference === '')
        data.loanReference =  getIncrementedLoanReference(getLastLoanReferenceOfEntity(data.entityName));
    var rowToInsert = buildLoanToInsert(data);
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var lastEntityRow = getLastLoanOfEntityRow(data.entityName);
    var rangeRowToSet = loansOriginalSheet.getRange(lastEntityRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.lastLoansColumn) + 1
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
    var interestRatePercent = data.interestRate / 100;
    row[ColumnNames.letterToColumnStart0('A')] = data.loanReference;
    row[ColumnNames.letterToColumnStart0('B')] = data.lenderReference;
    row[ColumnNames.letterToColumnStart0('C')] = data.entityName;
    row[ColumnNames.letterToColumnStart0('D')] = data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('E')] = data.dateBorrowed;
    row[ColumnNames.letterToColumnStart0('F')] = '';
    row[ColumnNames.letterToColumnStart0('G')] = data.dueDate;
    row[ColumnNames.letterToColumnStart0('H')] = interestRatePercent;
    row[ColumnNames.letterToColumnStart0('I')] = interestRatePercent * data.amountBorrowed;
    row[ColumnNames.letterToColumnStart0('J')] = 'No';
    row[ColumnNames.letterToColumnStart0('K')] = data.ballooninvestment;
    row[ColumnNames.letterToColumnStart0('L')] = '';
    row[ColumnNames.letterToColumnStart0('M')] = data.borrowerEntity;
    row[ColumnNames.letterToColumnStart0('N')] = '';

    return row;
}


function getLastLoanReferenceOfEntity(entityName) {
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var lastLoanOfEntityRow = getLastLoanOfEntityRow(entityName);
    return loansOriginalSheet.getRange(lastLoanOfEntityRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn)).getValue();
}

function getLastLoanOfEntityRow(entityName) {
    var lastRow = -1;
    var allLoans = getAllLoans();
    var loanReference = getLastLoanReferenceOfEntity(entityName);
    if(loanReference !== null) { // A loan of this entity has already been imported
        for(var i=0; i < allLoans.length; i++){
            var currentLoanReference = allLoans[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn)];
            if( currentLoanReference === loanReference)
                lastRow = i + INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow;
        }
        return lastRow;
    }
    else { // First loan of this entity to be imported
       /* var beforeEntityLoan = getLastLoanOfEntityBeforeThisEntity(entityName);
        if(beforeEntityLoan !== null) {
            var beforeEntityName = beforeEntityLoan[ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn)];
            for (var i = 0; i < allLoans.length; i++) {
                var currentLoanEntityName = allLoans[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn)];
                if (currentLoanEntityName === beforeEntityName)
                    lastRow = i + INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow;
            }
            return lastRow;
        }
        else // First loan of this entity to be imported and no entity with a name before this one in the list of loans
            return INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow - 1;*/
            
        // Add the loan at the end of the loan list (several empty lines at the end of the loans list need to be ignoredà
        var row = 0;
        while (row < allLoans.length && allLoans[row][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn)] !== '')
            row++;
        return row + 2;  // 2 Because of the two header lines
    }
}

function getAllLoans() {
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var loansRange = loansOriginalSheet.getRange(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        loansOriginalSheet.getLastRow(),
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn) -
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn)+1);
    return loansRange.getValues();
}

function getLastLoanOfEntityBeforeThisEntity(entityName) {
    var allLoans = getAllLoans();
    var entityNameColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn);
    var retVal = null;
    for (var i = 0; i < allLoans.length; i++) {
        var currentLoan = allLoans[i];
        var currentLoanName = currentLoan[entityNameColS0];
        if(currentLoanName !== "" && currentLoanName.localeCompare(entityName) < 0 )
            retVal = currentLoan;
        else
            return retVal;
    }
    return retVal;
}

function getLastLoanReferenceOfEntity(entityName){
    var allLoans = getAllLoans();
    var entityNameColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn);
    var loanReferenceColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.loanReferenceColumn);
    allLoans = allLoans.filter(function (loan) {
        return loan[entityNameColS0] === entityName;
    });
    if(allLoans.length === 0) // First loan of this entity to be imported
        return null;
    allLoans.sort(function (a, b) {
        return b[loanReferenceColS0].localeCompare([loanReferenceColS0]);
    });
    return allLoans[allLoans.length-1][loanReferenceColS0];
}

function getIncrementedLoanReference(previousLoanReference) {
    if(previousLoanReference === null) // The loan to be created is the first loan of this entity, hence there is no loan reference in the sheet yet
        return "LOAN001";
    // Split in two strings: letters, and digits (loan references
    // are a concatenation of a group of letters and a group of numbers
    var splittedOldLoanReference = (/([A-Z]+)([0-9]{3}).*/g).exec(previousLoanReference);
    var oldLoanNumberStr = splittedOldLoanReference[2];
    var oldLoanNumber = parseInt(oldLoanNumberStr, 10);
    var incrementedLoanNumber = oldLoanNumber+1;
    if(incrementedLoanNumber > 999)
        incrementedLoanNumber = 999
    var incrementedLoanNumberStr = ""+incrementedLoanNumber;
    Logger.log(incrementedLoanNumberStr);
    while (incrementedLoanNumberStr.length < 3) {
        incrementedLoanNumberStr = "0" + incrementedLoanNumberStr;
    }
    return splittedOldLoanReference[1] + incrementedLoanNumberStr;
}