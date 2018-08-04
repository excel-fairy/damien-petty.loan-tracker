var INTEREST_STATEMENT_SHEET = {
    name: 'Interest statement',
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interest statement'),
    dateCell: 'H1',
    pdfExportRange: {
        r1: 5,
        r2: 47,
        c1: 1,
        c2: 8
    }
};
