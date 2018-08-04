var ENTITIES_SHEET = {
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entity"),
    entityNameColumn: letterToColumnStart0('A'),
    emailAddressColumn: letterToColumnStart0('G'),
    emailSubjectColumn: letterToColumnStart0('M'),
    emailBodyColumn: letterToColumnStart0('N'),
    entitiesListRange:{
        r1: 3,
        r2: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entity").getLastRow(),
        c1: letterToColumn('A'),
        c2: letterToColumn('N')
    }
};

var LOANS_SHEET = {
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loans")
};

var TEST_INTEREST_SHEET = {
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test_Interest")
};


function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Manage loans')
        .addItem('Import', 'openCreateLoanPopup')
        .addToUi();
}


// sample usage:
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




function createLoan(data) {
    SpreadsheetApp.getUi().alert ("Loan is being imported. Please wait for it to be fully created");
    appendLoanToLoansSheet(data);
    appendTestInterests(data);
}

function appendLoanToLoansSheet(data){
    var row = [];
    row[letterToColumnStart0('A')] = 'TODO'; //TODO
    row[letterToColumnStart0('B')] = '';
    row[letterToColumnStart0('C')] = data.entityName;
    row[letterToColumnStart0('D')] = data.amountBorrowed;
    row[letterToColumnStart0('E')] = data.dateBorrowed;
    row[letterToColumnStart0('F')] = '☐';
    row[letterToColumnStart0('G')] = data.dueDate;
    row[letterToColumnStart0('H')] = data.interestRate;
    row[letterToColumnStart0('I')] = data.interestRate * data.amountBorrowed;
    row[letterToColumnStart0('J')] = 'No';
    row[letterToColumnStart0('K')] = '';
    row[letterToColumnStart0('L')] = data.borrowerEntity;
    LOANS_SHEET.sheet.appendRow(row); //TODO: position of appended row
}

function appendTestInterests(data){
    for (var i = 0; i < 12; i++){
        var date = new Date();  //TODO: Current year ?
        date.setDate(1);
        date.setMonth(i);
        var row = [];
        var nbDaysInMonth = getNbDaysInMonth(date.getMonth(), date.getFullYear());
        row[letterToColumnStart0('A')] = date; //TODO: formatting
        row[letterToColumnStart0('B')] = nbDaysInMonth;
        row[letterToColumnStart0('C')] = nbDaysInMonth;
        row[letterToColumnStart0('D')] = 'TODO'; //TODO: Same as column A in Loans sheet
        row[letterToColumnStart0('E')] = data.interestRate;
        row[letterToColumnStart0('F')] = data.entityName;
        row[letterToColumnStart0('G')] = data.amountBorrowed;
        row[letterToColumnStart0('H')] = data.amountBorrowed * (data.interestRate / nbDaysInMonth); //TODO: Check helene's answer
        row[letterToColumnStart0('I')] = '';
        row[letterToColumnStart0('J')] = '';
        row[letterToColumnStart0('K')] = '';
        row[letterToColumnStart0('L')] = '';
        row[letterToColumnStart0('M')] = '';
        row[letterToColumnStart0('N')] = getMonthFullName(date.getMonth()+1);
        row[letterToColumnStart0('O')] = date.getFullYear();
        row[letterToColumnStart0('P')] = '';
        row[letterToColumnStart0('Q')] = '';
        row[letterToColumnStart0('R')] = '';
        row[letterToColumnStart0('S')] = '';
        row[letterToColumnStart0('T')] = '';
        row[letterToColumnStart0('U')] = '';
        row[letterToColumnStart0('V')] = '';
        row[letterToColumnStart0('W')] = '';
        row[letterToColumnStart0('X')] = '';
        TEST_INTEREST_SHEET.sheet.appendRow(row);
    }
}

function getNbDaysInMonth (month, year) {
    return new Date(year, month, 0).getDate();
}

// Already exists
function getEntitiesNames(){
    var entities = ENTITIES_SHEET.sheet.getRange(ENTITIES_SHEET.entitiesListRange.r1,
        ENTITIES_SHEET.entitiesListRange.c1,
        ENTITIES_SHEET.entitiesListRange.r2 - ENTITIES_SHEET.entitiesListRange.r1,
        ENTITIES_SHEET.entitiesListRange.c2 - ENTITIES_SHEET.entitiesListRange.c1+1).getValues();
    return entities.map(function(entity){return entity[ENTITIES_SHEET.entityNameColumn];});
}

function getMonthFullName(month){
    var months = [
        'January', 'February', 'March', 'April', 'May', 'June', 'July',
        'August', 'September', 'October', 'November', 'December'
    ];
    return months[month-1];
}