function onOpen() {
  DocumentApp.getUi()
    .createMenu("Contract Generator")
    .addItem("Create New Contract", "showContractForm")
    .addToUi();
}

function showContractForm() {
  var html = HtmlService.createHtmlOutputFromFile("ContractForm")
    .setWidth(400)
    .setHeight(600);
  DocumentApp.getUi().showModalDialog(html, "Create New Contract");
}

function processForm(formData) {
  try {
    Logger.log("Form Data Received: " + JSON.stringify(formData));

    var emailAddress = formData.emailAddress;
    delete formData.emailAddress;

    // Parse the date components from the form data
    var dateParts = formData.date.split("-");
    var year = parseInt(dateParts[0], 10);
    var month = parseInt(dateParts[1], 10) - 1; // Months are zero-based in JavaScript
    var day = parseInt(dateParts[2], 10);

    // Create a Date object without time zone offset
    var date = new Date(year, month, day);

    // Format the date using Utilities.formatDate with 'GMT' time zone
    var formattedDate = Utilities.formatDate(date, "GMT", "MMMM d, yyyy");

    // Sanitize the amount input to remove any non-numeric characters except the decimal point
    var amountValue = formData.amount.replace(/[^0-9.]/g, "");

    // Parse the sanitized amount and format it
    var formattedAmount = "$" + parseFloat(amountValue).toFixed(2);

    var clientData = {
      "{{Date}}": formattedDate,
      "{{CompanyName}}": "Kickstand Services",
      "{{ClientName}}": formData.clientName,
      "{{ServiceDescription}}": formData.serviceDescription,
      "{{Amount}}": formattedAmount,
      "{{RepresentativeName}}": formData.representativeName,
    };

    Logger.log("Client Data: " + JSON.stringify(clientData));

    // Generate the contract and get the PDF file
    var pdfFile = generateContract(clientData);

    // Email the PDF
    var subject = "Your Contract with " + clientData["{{CompanyName}}"];
    var message =
      "Dear " +
      clientData["{{ClientName}}"] +
      ",\n\nPlease find attached your contract.\n\nBest regards,\n" +
      clientData["{{RepresentativeName}}"];

    MailApp.sendEmail(emailAddress, subject, message, {
      attachments: [pdfFile],
    });

    // Log the data to Google Sheets
    logContractData(formData, pdfFile.getUrl());

    return "Contract generated, emailed, and logged successfully.";
  } catch (e) {
    Logger.log("Error in processForm: " + e.message);
    throw new Error("Failed to process form: " + e.message);
  }
}

function generateContract(clientData) {
  var TEMPLATE_ID = "1JjB-OYxBeW0ig9B35EzUzX2xjZSdnO03dpO3OjLX1WQ"; // Replace with your Template Document ID
  var DEST_FOLDER_ID = "1uyqc6XooPaNPBDFQpI3D2bZPSlT1rBOm"; // Replace with your Destination Folder ID

  try {
    Logger.log("Template ID: " + TEMPLATE_ID);

    // Open the template document
    var templateDoc = DriveApp.getFileById(TEMPLATE_ID);

    // Make a copy of the template document
    var newDocFile = templateDoc.makeCopy(
      "Contract - " + clientData["{{ClientName}}"]
    );
    var newDocId = newDocFile.getId();
    Logger.log("New Document ID: " + newDocId);

    // Open the new document for editing
    var newDoc = DocumentApp.openById(newDocId);
    var newDocBody = newDoc.getBody();

    // Replace placeholders with actual data
    for (var placeholder in clientData) {
      newDocBody.replaceText(placeholder, clientData[placeholder]);
    }

    // Save and close the new document
    newDoc.saveAndClose();

    // Pause to ensure changes are saved before conversion
    Utilities.sleep(2000); // Wait for 2 seconds

    // Convert the document to PDF
    var pdfFile = DriveApp.getFileById(newDocId).getAs("application/pdf");
    pdfFile.setName("Contract - " + clientData["{{ClientName}}"] + ".pdf");

    // Save the PDF to the destination folder
    var destFolder = DriveApp.getFolderById(DEST_FOLDER_ID);
    var savedPdf = destFolder.createFile(pdfFile);

    // Optionally, delete the temporary Google Doc
    DriveApp.getFileById(newDocId).setTrashed(true);

    Logger.log("Contract generated and saved as PDF.");

    // Return the saved PDF file
    return savedPdf;
  } catch (e) {
    Logger.log("Error in generateContract: " + e.message);
    throw new Error("Failed to generate contract: " + e.message);
  }
}

function logContractData(formData, pdfUrl) {
  var SPREADSHEET_ID = "1OwgYjtsS53iGHhgxYmxuv0TU-HO3N-8LqcozazZhCL4"; // Replace with your Spreadsheet ID
  var SHEET_NAME = "Contracts"; // The name of the sheet where you want to log the data

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    // If the sheet doesn't exist, create it and set up headers
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Timestamp",
        "Date",
        "Client Name",
        "Service Description",
        "Amount",
        "Representative Name",
        "Client Email",
        "Contract PDF URL",
      ]);
    }

    // Get the current timestamp
    var timestamp = new Date();

    // Append the data to the sheet
    sheet.appendRow([
      timestamp,
      formData.date,
      formData.clientName,
      formData.serviceDescription,
      formData.amount,
      formData.representativeName,
      formData.emailAddress,
      pdfUrl,
    ]);

    Logger.log("Contract data logged to spreadsheet.");
  } catch (e) {
    Logger.log("Error in logContractData: " + e.message);
    throw new Error("Failed to log contract data: " + e.message);
  }
}
