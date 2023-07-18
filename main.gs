// gmail2sheet
//
// TODO: Set up a trigger to run main()

//
// Configurable constants.
//
const kSpreadsheetId = "1JignDYAjNdLCyieHHJUrhe3zR6d4n6pIkvLhRY45jHs";
const kGmailQueryBase = "label:tmp-junk";

// Constants. Don't change these.
const kGmailProcessedLabel = "processed-by-gmail2sheet";

function readPersistentStorage(sheet) {
  let users_range = sheet.getRange(2, 2, 1, 1);
  let num_users = users_range.getCell(1, 1).getValue();
  let users_rows = sheet.getRange(3, 2, num_users, 1).getValues();
  let users = [];
  for (let row of users_rows) {
    users.push(row[0]);
  }

  let info = {
    users: users,
  };
  return info;
}

function writePersistentStorage(info, sheet) {
  sheet.getRange(1,1).getCell(1, 1).setValue(info.lastRowWritten);
}

function writeGmailThreadToRow(thread, sheet, row) {
  let subject = thread.getFirstMessageSubject();
  sheet.getRange(1,1).getCell(1,1).setValue(subject);
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

  let gmail_query = kGmailQueryBase + " -label:" + kGmailProcessedLabel;
  let threads = GmailApp.search(gmail_query);

  let spreadsheet = SpreadsheetApp.openById(kSpreadsheetId);
  // The storage sheet is used to persist data between runs. The export sheet is
  // where we write new messages to.
  let storage_sheet = spreadsheet.getSheetByName("Storage");
  let export_sheet = spreadsheet.getSheetByName("Export");

  let storage = readPersistentStorage(storage_sheet);

  // Change color to indicate that the script is running.

  storage_sheet.setTabColor('red');
  export_sheet.setTabColor('red');

  if (threads.length > 0) {
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
    const kNumColumns = 3;
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
      ];
      rows.push(row);
    }

    // Make space and write the new rows.
    export_sheet.insertRowsBefore(1, rows.length);
    const out_range = export_sheet.getRange(1, 1, rows.length, kNumColumns);
    out_range.setValues(rows);

    processed_label.addToThreads(threads);
  }

  storage_sheet.setTabColor('blue');
  export_sheet.setTabColor('blue');
}
