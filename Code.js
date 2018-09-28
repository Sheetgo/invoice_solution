/*================================================================================================================*
  Invoicing Solution by Sheetgo
  ================================================================================================================
  Version:      1.0.0
  Project Page: https://github.com/Sheetgo/supplier_system
  Copyright:    (c) 2018 by Sheetgo
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  ----------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.0  Initial release
 *================================================================================================================*/


/**
 * Template file id and names.
 * This configuration changes after the script copy the template files
 * @type
 */
Files = {
    // Supplier
    Form_Supplier: { type: "form", id: null, name: "Supplier Registration Form" },
    Ss_Supplier_Database: { type: "spreadsheet", id: "1sNQK4ceopSMj-JC5-MmIZ2RZN4dVHZC9VatQzbSB-40", name: "Supplier Database" },

    // Invoice
    Form_Invoice: { type: "form", id: null, name: "Invoices Registration Form" },
    Ss_Invoice_Database: { type: "spreadsheet", id: "13skhyt9AB29oGF6v4DHxHhhEfGTI8na9O8eqtv6WurA", name: "Invoice Database" },

    // Dashboard
    Ss_Invoice_Dashboard: { type: "spreadsheet", id: null, name: "Suppliers Invoices Dashboard" }
}


/**
 * Creates the 'Suppliers' Menu in the spreadsheet. This function is fired every time a spreadsheet is open
 * @param {JSON} e User/Spreadsheet basic parameters 
 */
function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    var menu = ui.createMenu('Suppliers')
    if (e && e.authMode == ScriptApp.AuthMode.LIMITED) {
        menu.addItem('Create Solution System', 'create_system')
    } else {
        menu.addItem('Send Payment Done Email', 'sendInvoiceEmail')
    }
    menu.addToUi();
}


/**
 * Create the Supplier Invoice system by copying the template files and moving into an local 
 * user folder within Google Drive
 */
function create_system() {

    // Dashboad Spreadheet (Main Spreadsheet) Object
    var ss_dashboard = SpreadsheetApp.getActiveSpreadsheet();
    Files.Ss_Invoice_Dashboard.id = ss_dashboard.getId();

    ss_dashboard.toast("Creating & Configuring Solution. Please wait...");

    // Create the Solution folder on users Drive 
    var folder = DriveApp.createFolder("Sheetgo Suppliers System");

    // Move the current Dashboard spreadsheet into the Solution folder
    var file = DriveApp.getFileById(Files.Ss_Invoice_Dashboard.id);
    file.setName(Files.Ss_Invoice_Dashboard.name);
    moveFile(file, folder);

    // Create Invoice Spreadsheet & Form
    create_spreadsheet_and_form(Files.Ss_Invoice_Database, Files.Form_Invoice, folder, "Creating Invoice Form & Spreadsheet...");

    // Create Supplier Spreadsheet & Form
    create_spreadsheet_and_form(Files.Ss_Supplier_Database, Files.Form_Supplier, folder, "Creating Supplier Form & Spreadsheet...");

    // Record the Invoice Registration Form ID on Supliers Database Spreadsheet
    var spreadsheet = SpreadsheetApp.openById(Files.Ss_Supplier_Database.id).getSheetByName("Settings");

    // TODO: Get the Form ID from the copyed forms
    spreadsheet.getRange("B1").setValue(Files.Form_Invoice.id);
    spreadsheet.getRange("B2").setValue(Files.Form_Supplier.id);

    // Update menu
    onOpen();
}   


/**
 * Create the spreadsheet from template, rename it and move it into the solution folder
 * @param {JSON} spreadsheet Spreadsheet dic information 
 * @param {JSON} form Form dic information 
 * @param {Object} folder Folder object in Google Drive 
 * @param {string} message Message to display to the user 
 */
function create_spreadsheet_and_form(spreadsheet, form, folder, message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
    
    // Create a spreadsheet copy from the template and move it into the solution folder. 
    // This process will also create a copy from the bounded Form
    var file = DriveApp.getFileById(spreadsheet.id).makeCopy(spreadsheet.name, folder);
    moveFile(file, folder);

    // Set the spreadsheet id on the variable
    spreadsheet.id = file.getId();

    // Get the id information from the new Form created, rename the file and move it into the solution folder
    form.id = FormApp.openByUrl(SpreadsheetApp.openById(spreadsheet.id).getFormUrl()).getId();
    var formFile = DriveApp.getFileById(form.id);
    formFile.setName(form.name); // Rename file
    moveFile(formFile, folder);
}


/**
 * Send the email notification after the payment is done
 */
function sendInvoiceEmail() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  // Get the spreadsheet
    var ss = spreadsheet.getSheetByName("Invoices paid");     // Get the "Invoice paid" tab
    var data = ss.getDataRange().getValues();                 // Get all data from invoice paid tab

    // Get email body and email subject from the "Email Data" tab
    var sheetEmailData = spreadsheet.getSheetByName("Email data");
    var emailText = sheetEmailData.getRange("C5").getValue();
    var emailSubject = sheetEmailData.getRange("C4").getValue();

    // Scan the array of invoices to look for email to be sent
    for (var i = 1; i < data.length; i++) {
        var row = data[i];        // Data from the current row
        var emailSent = row[12];  // Indicates if the email has already been sent
        var name = row[1];        // Name of the person that received the payment
        var recipient = row[0];   // Email of the person that received the payment
        var currency = row[4];    // Currency used in the invoice/payment
        var amount = row[3];      // Invoice amout

        /* Send the email */
        if (row[0] && !emailSent) {
            var body = emailText.replace("%name%", name).replace("%currency%", currency).replace("%amount%", amount)
            MailApp.sendEmail({
                to: recipient,
                subject: emailSubject,
                htmlBody: body
            });

            // Mark as email sent in the tab "Invoices paid"
            ss.getRange("M" + (i + 1)).setValue(true);
        }
    }
}


/**
 * Move a file from one folder into another
 * @param {Object} file A file object in Google Drive
 * @param {Object} dest_folder A folder object in Google Drive 
 */
function moveFile(file, dest_folder) {
    dest_folder.addFile(file);
    var parents = file.getParents();
    while (parents.hasNext()) {
        var folder = parents.next();
        if (folder.getId() != dest_folder.getId()) {
            folder.removeFile(file)
        }
    }
}