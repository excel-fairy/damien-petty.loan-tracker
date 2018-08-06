function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Run scripts')
        .addItem('Import loan', 'openCreateLoanPopup')
        .addItem('Export all interest statements and send via email (slow)', 'exportInterestStatements')
        .addItem('Export invoices as CSV', 'exportInvoicesAsCsv')
        .addToUi();
}
