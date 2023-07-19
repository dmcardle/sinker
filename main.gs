// gmail2sheets
//
// This script exports parts of matching messages from Gmail to Sheets. It is
// intended to be run repeatedly, so it only picks up new threads each time. It
// marks threads that have already been exported by adding a Gmail label.
//
// TODO: Install menu button to enable users to run this script manually?
// TODO: Alternatively, figure out how to set up a repeating trigger.
// TODO: Automatically assign users from rotation?
// TODO: Automatically detect status of each thread by checking whether a
//       rotation member has responded?

//---------------------------------------------------------------------------
//                             Configurable constants
//---------------------------------------------------------------------------
// This is the name of the Gmail label used to mark threads that have already
// been exported. Customize it to describe your use case. If you have multiple
// instances of this script, it's critical that this value is unique.
const kGmailProcessedLabel = "gmail2sheets/processed/JobAlerts";
// This is the query used to match emails from your Gmail account.
const kGmailQueryBase = "from:(Job Alerts from Google)";
// Get this value from the URL of the desired spreadsheet.
const kSpreadsheetId = "1JignDYAjNdLCyieHHJUrhe3zR6d4n6pIkvLhRY45jHs";
//---------------------------------------------------------------------------

// Parses data from the "Storage" sheet into a nice object.
function readPersistentStorage(sheet) {
  let users = sheet
      .getRange("A2:A")
      .getValues()
      .map(row => row[0])
      .filter(s => s.length > 0);
  if (users.length == 0) {
    throw new Error("There must be more than zero users in the rotation.");
  }
  let info = {
    users: users,
  };
  return info;
}

function writePersistentStorage(data, sheet) {
  sheet.getRange("A1").setValue("Rotation Members");
  let users_range = sheet.getRange("A2:A");
  users_range = users_range.offset(0, 0, data.users.length, 1);
  users_range.setValues(data.users.map(s => [s]));
}

function writeGmailThreadToRow(thread, sheet, row) {
  let subject = thread.getFirstMessageSubject();
  sheet.getRange(1,1).getCell(1,1).setValue(subject);
}

// Gets a Sheet from the given Spreadsheet. If it doesn't exist, it creates a
// sheet with the given name.
function getOrCreateSheet(spreadsheet, sheet_name) {
  let sheet = spreadsheet.getSheetByName(sheet_name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheet_name);
  }
  if (!sheet) {
    throw new Error("Failed to create new sheet: " + sheet_name)
  }
  return sheet;
}

function main() {
  // Get or create the "processed" label.
  let processed_label = GmailApp.getUserLabelByName(kGmailProcessedLabel);
  if (!processed_label) {
    processed_label = GmailApp.createLabel(kGmailProcessedLabel);
  }
  if (!processed_label) {
    throw new Error("Failed to get label: " + kGmailProcessedLabel);
  }

  // Open the spreadsheet and get handles to the sheets we need. The storage
  // sheet is used to persist data between runs. The export sheet is where
  // messages are exported.
  let spreadsheet = SpreadsheetApp.openById(kSpreadsheetId);
  let storage_sheet = getOrCreateSheet(spreadsheet, "Storage");
  let export_sheet = getOrCreateSheet(spreadsheet, "Export");

  // Change tab color to indicate that the script is running.
  storage_sheet.setTabColor('red');
  export_sheet.setTabColor('red');

  let storage = readPersistentStorage(storage_sheet);
  let gmail_query = kGmailQueryBase + " -label:" + kGmailProcessedLabel;
  let threads = GmailApp.search(gmail_query);
  exportThreads(threads, processed_label, export_sheet);
  writePersistentStorage(storage, storage_sheet);

  storage_sheet.setTabColor('blue');
  export_sheet.setTabColor('blue');
}

function exportThreads(threads, processed_label, export_sheet) {
  if (threads.length == 0) {
    return;
  }

  // Sort threads by date.
  threads.sort((a,b) => {
    if (a.getLastMessageDate() < b.getLastMessageDate())
      return 1;
    else if (a.getLastMessageDate() > b.getLastMessageDate())
      return -1;
    else
      return 0;
  });

  // Convert threads into spreadsheet rows.
  const rows = [];
  for (const thread of threads) {
    const messages = thread.getMessages();
    if (messages.length == 0) {
      throw new Error("Encountered impossible thread with zero messages.");
    }
    const first_message = messages[0];
    const row = [
      first_message.getFrom(),
      first_message.getSubject(),
      first_message.getDate(),
      first_message.getId(),
    ];
    rows.push(row);
  }
  let num_columns = rows[0].length;

  // Make space and write the new rows.
  export_sheet.insertRowsBefore(1, rows.length);
  const out_range = export_sheet.getRange(1, 1, rows.length, num_columns);
  out_range.setValues(rows);

  processed_label.addToThreads(threads);
}
