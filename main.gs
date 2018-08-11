var INTEREST_STATEMENT_SPREADSHEET = {
    entitiesSheet: { // ImportRange from Loan tracker spradsheet
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entity"),
        entityNameColumn: ColumnNames.letterToColumnStart0('A'),
        emailAddressColumn: ColumnNames.letterToColumnStart0('G'),
        emailSubjectColumn: ColumnNames.letterToColumnStart0('M'),
        emailBodyColumn: ColumnNames.letterToColumnStart0('N'),
        carbonCopyEmailAddressesColumn: ColumnNames.letterToColumnStart0('O'),
        entitiesListRange:{
            r1: 3,
            r2: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entity").getLastRow(),
            c1: ColumnNames.letterToColumn('A'),
            c2: ColumnNames.letterToColumn('O')
        }
    },
    interestStatementSheet: {
        name: 'Interest statement',
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interest statement'),
        dateCell: 'H1',
        totalCell: 'H35',
        entityCell: 'C3',
        pdfExportRange: {
            r1: 5,
            r2: 47,
            c1: 1,
            c2: 8
        }
    },
    invoicesSheet: {
        name: 'Invoices',
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CSV file export"),
        descriptionColumn: ColumnNames.letterToColumnStart0('Q'),
        invoiceNumberColumn: ColumnNames.letterToColumnStart0('K'),
        exportRange:{
            r1: 1,
            r2: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CSV file export").getLastRow(),
            c1: ColumnNames.letterToColumn('A'),
            c2: ColumnNames.letterToColumn('AA')
        }
    },
    calcSheet: {
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calc"),
        lastInvoiceNumberCell: 'I2'
    },
    loansSheet: {
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loans"),
        name: 'Loans',
        firstLoanRow: '2',
        firstLoansColumn: 'A',
        lastLoansColumn: 'M',
        loanReferenceColumn: 'A',
        entityNameColumn: 'C'
    }
};


function getEntitiesNames(){
    var entities = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1+1).getValues();
    return entities.map(function(entity){return entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entityNameColumn];});
}

