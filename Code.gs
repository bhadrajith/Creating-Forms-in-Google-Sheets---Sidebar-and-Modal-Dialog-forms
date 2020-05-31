//CREATE CUSTOM MENU
function onOpen() { 
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("My Menu")
    .addItem("Sidebar Form","showFormInSidebar")
    .addItem("Modal Dialog Form","showFormInModalDialog")
    .addItem("Modeless Dialog Form","showFormInModlessDialog")
    .addToUi();
}

//OPEN THE FORM IN SIDEBAR 
function showFormInSidebar() {      
  var form = HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Contact Details');
  SpreadsheetApp.getUi().showSidebar(form);
}


//OPEN THE FORM IN MODAL DIALOG
function showFormInModalDialog() {
  var form = HtmlService.createTemplateFromFile('Index').evaluate();
  SpreadsheetApp.getUi().showModalDialog(form, "Contact Details");
}


//OPEN THE FORM IN MODALLESS DIALOG
function showFormInModlessDialog() {
  var form = HtmlService.createTemplateFromFile('Index').evaluate();
  SpreadsheetApp.getUi().showModelessDialog(form, "Contact Details");
}


//PROCESS FORM
function processForm(formObject){ 
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([formObject.first_name,
                formObject.last_name,
                formObject.gender,
                formObject.email
                //Add your new field names here
                ]);
}

//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
