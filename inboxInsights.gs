// Adds a custom menu item "Get Email Summary" to the sheet

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Summary')
    .addItem('Get Email Summary', 'numberOfEmails')
    .addToUi();
}

// Prompts the user to input the number of emails to be fetched
function numberOfEmails() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Email Summary', 'Enter the number of emails to fetch:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const numberOfEmails = parseInt(response.getResponseText());
    getGmailThreads(numberOfEmails);
  }
}

// Retreives each email thread and passes the required data into an array
function getGmailThreads(numberOfEmails) {
  const gmailThreads = GmailApp.getInboxThreads()
  let i = 0;
  let emailDetails = []

  for (i = 0; i < numberOfEmails; i++) {
    const thread = gmailThreads[i];
    const messages = thread.getMessages();
    let mailBody = messages[0].getPlainBody();
    let summary = summarizeBody(mailBody);
    let mailSubject = messages[0].getSubject();
    let mailDate = messages[0].getDate();
    let senderDetails = separateSenderDetails(messages[0].getFrom());
    emailDetails.push([
      senderDetails.name,
      senderDetails.email,
      mailDate,
      mailSubject,
      summary
    ]);
  }
  displayDetails(emailDetails);
}

// Displays the processed data in the Google Sheet
function displayDetails(emailDetails) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.setName("Email Summary");
  const headers = ["Sender Name", "Sender Email", "Date", "Subject", "Body Preview"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, emailDetails.length, emailDetails[0].length).setValues(emailDetails);

  formatSheet(sheet, emailDetails.length + 1);

  SpreadsheetApp.getUi().alert('Email summary fetched successfully!');
}

//Formats the Google Sheet to enhance readability 
function formatSheet(sheet, numberOfRows) {

  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#6C5B7B');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  const bandingRange = sheet.getRange(2, 1, numberOfRows - 1, 5);
  const banding = bandingRange.applyRowBanding();
  banding.setFirstRowColor('#FFFFFF');
  banding.setSecondRowColor('#F0F0F0');
  const senderNameColumn = sheet.getRange(2, 1, numberOfRows, 1);
  senderNameColumn
    .setFontWeight('bold')
  const senderEmailColumn = sheet.getRange(2, 2, numberOfRows, 1)
  senderEmailColumn
    .setFontStyle('italic');
  const fullRange = sheet.getRange(1, 1, numberOfRows, 5);
  fullRange.setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 5);
  sheet.setFrozenRows(1);
}

//Summarizes the email body using OpenAI's API 
function summarizeBody(emailBody) {

  const GPT_API = "{API_KEY}";
  const BASE_URL = "https://api.openai.com/v1/chat/completions";

  const headers = {
    "Content-Type": "application/json",
    "Authorization": `Bearer ${GPT_API}`
  };


  const options = {
    headers,
    method: "GET",
    muteHttpExceptions: true,
    payload: JSON.stringify({
      "model": "gpt-4o",
      "messages": [{
        "role": "system",
        "content": "Read the body of the email and summarize the email in less than 80 words. The summary must provide a gist of the entire email provided."
      },
      {
        "role": "user",
        "content": emailBody
      }
      ],
      "temperature": 0.5
    })
  }
  const response = JSON.parse(UrlFetchApp.fetch(BASE_URL, options));
  return response.choices[0].message.content;
}


//Seperates the sender name and email 
function separateSenderDetails(senderDetails) {
  let startEmail = senderDetails.indexOf('<');
  let endEmail = senderDetails.indexOf('>')
  let sender = {
    name: "",
    email: ""
  }
  if (startEmail > 0) {
    sender.name = senderDetails.substring(0, startEmail).trim();
    sender.email = senderDetails.substring(startEmail + 1, endEmail);
  }
  else {
    sender.name = "Not Available";
    sender.email = senderDetails;
  }
  return sender;
}
