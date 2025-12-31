function onEdit(e) {

  // 1. BASIC VALIDATION
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();
  
  // MODIFICATION: Check if the sheet name is exactly "Round 2 Phase 1"
  if (sheet.getName() == "Round 2 Phase 1" || sheet.getName() == "Round 2 Phase 2")
  {

  var row = e.range.getRow();
  var col = e.range.getColumn();
  var value = e.value; // The new value entered
  
  // Configuration
  var ENCODING_START_ROW = 2;
  var ENCODING_END_ROW = 30;
  var DECODING_START_ROW = 32;
  var DECODING_END_ROW = 62;
  
  // Check if edit is within Encoding Team Rows
  if (row < ENCODING_START_ROW || row > ENCODING_END_ROW) return;
  
  // Check if column is a Cipher Column (Starts at D=4, F=6... must be even)
  if (col < 4 || col % 2 !== 0) return;
  
  // Check if value is "Y" (case insensitive)
  if (!value || value.toString().toUpperCase() !== "Y") return;
  if (sheet.getRange(row, col + 1).getValue() !== "") return;

  // LOCK SERVICE: Prevents race conditions if two proctors edit simultaneously
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait up to 10 seconds for other scripts to finish
  } catch (e) {
    SpreadsheetApp.getUi().alert('System is busy. Please try entering "Y" again.');
    return;
  }

  // 2. GENERATE ID
  // Get Encoding Team Details
  // Col A is index 1, Col B is index 2
  var encodingTeamId = sheet.getRange(row, 1).getValue(); 
  var encodingTeamName = sheet.getRange(row, 2).getValue();
  
  // Calculate Cipher Number: (Col 4 is Cipher 1, Col 6 is Cipher 2) -> (Col - 2) / 2
  var cipherNum = (col - 2) / 2;
  
  var uniqueId = String(encodingTeamId) + String(cipherNum);
  
  // Update the cell immediately to the Unique ID
  e.range.setValue(uniqueId);
  
  // 3. FIND DECODING TEAM (LOAD BALANCING & LEFTMOST SLOT)
  
  // Get all decoding data in one batch to be fast
  var totalCols = sheet.getLastColumn();
  var decodingRange = sheet.getRange(DECODING_START_ROW, 1, (DECODING_END_ROW - DECODING_START_ROW + 1), totalCols);
  var decodingData = decodingRange.getValues();
  
  var eligibleTeams = [];
  var minWorkload = 9999;

  // Loop through every decoding team
  for (var i = 0; i < decodingData.length; i++) {
    var teamRowData = decodingData[i];
    var teamId = teamRowData[0]; // Col A
    
    if (!teamId) continue; // Skip empty rows

    // Calculate Workload and Find First Empty Slot
    var currentWorkload = 0;
    var firstEmptyCol = -1;

    // Check Cipher columns: D(idx 3), F(idx 5), H(idx 7)...
    for (var k = 3; k < teamRowData.length; k += 2) { 
        // Check if slot has a cipher assigned
        if (teamRowData[k] !== "" && teamRowData[k] !== null) {
            currentWorkload++;
        } else if (firstEmptyCol === -1) {
            // If cell is empty and we haven't found a spot yet, this is the leftmost empty column
            firstEmptyCol = k + 1; // Convert 0-index array to 1-based Sheet Column
        }
    }

    // A team is eligible ONLY if they have an empty slot available
    if (firstEmptyCol !== -1) {
      eligibleTeams.push({
        rowIndex: i,
        teamId: teamId,
        workload: currentWorkload,
        targetCol: firstEmptyCol // Store the specific column to write to for this team
      });

      // Track minimum workload found so far
      if (currentWorkload < minWorkload) {
        minWorkload = currentWorkload;
      }
    }
  }

  // 4. SELECT THE WINNING TEAM
  if (eligibleTeams.length === 0) {
    SpreadsheetApp.getUi().alert("Error: No decoding teams have space available.");
    lock.releaseLock();
    return;
  }

  // Filter for only the teams that have the MINIMUM workload (Load Balancing)
  var bestCandidates = eligibleTeams.filter(function(t) {
    return t.workload === minWorkload;
  });

  // Randomly select one from the best candidates
  var winnerIndex = Math.floor(Math.random() * bestCandidates.length);
  var selectedTeam = bestCandidates[winnerIndex];

  // 5. WRITE DATA TO SHEET
  
  // A. Update Encoding Row: "Pass To" column (Current Col + 1)
  sheet.getRange(row, col + 1).setValue(selectedTeam.teamId);

  // B. Update Decoding Row: 
  // We need to map the array index back to the sheet row number
  var targetSheetRow = DECODING_START_ROW + selectedTeam.rowIndex;
  var targetSheetCol = selectedTeam.targetCol; // Use the calculated leftmost empty column
  
  // Write Cipher ID in the target Cipher Column
  sheet.getRange(targetSheetRow, targetSheetCol).setValue(uniqueId);
  
  // Write Encoding Team Name in the "From" Column (Next to it)
  sheet.getRange(targetSheetRow, targetSheetCol + 1).setValue(encodingTeamName);

  // Release lock
  lock.releaseLock();

  //END ROUND 2 phases distro
  }

  var ss        = e.source;
  var sheet     = e.range.getSheet();
  var row       = e.range.getRow();
  var col       = e.range.getColumn();
  var name      = sheet.getName();
  var editValue = e.value;  // single‐cell edit

  // 1) Any edit in C2:C∞ → "P"
  if (col === 3 && row >= 2 && editValue !== undefined) {
    sheet.getRange(row, 3).setValue("P");
    return;
  }

  // 2) Edit in I2 → clear I2 & fill blanks in C2:C∞ with "A"
  if (col === 9 && row === 2) {
    sheet.getRange("I2").clearContent();
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var cRange = sheet.getRange(2, 3, lastRow - 1, 1);
      var vals   = cRange.getValues();
      for (var i = 0; i < vals.length; i++) {
        if (!vals[i][0]) vals[i][0] = "A";
      }
      cRange.setValues(vals);
    }
    return;
  }
}
/**
 * Polls Gmail every 2 minutes for new "submission_*" emails,
 * parses their subject, and writes a link into the appropriate sheet.
 */
function pollSubmissions() {
  // 1) Find unread threads with subjects starting "submission_"
  var threads = GmailApp.search('label:inbox is:unread subject:"submission"');
  var ss = SpreadsheetApp.getActive();
  var errors = ss.getSheetByName('Errors');
  
  threads.forEach(function(thread) {
    // Process each message in the thread
    thread.getMessages().forEach(function(msg) {
      if (!msg.isUnread()) return;

      var subj = msg.getSubject() || '';
      var parts = subj.split(' ');
      var roll, team, roundNum;
      
      if (parts.length === 4 && parts[0].toLowerCase() === 'submission') {
        roll = parts[1];
        team = parts[2];
        roundNum = parts[3];

      } else {
        // Unparseable subject → record an error row
        errors.appendRow([
          parts[1] || 'Unknown',
          parts[2] || 'Unknown',
          parts[3] || 'Unknown',
          getThreadLink(thread)
        ]);
        msg.markRead();
        return;
      }
      
      // Target sheet name
      var sheetName = 'Round ' + roundNum;
      var sheet = ss.getSheetByName(sheetName);

      var placed = false;
      if (sheet) {
        // Fetch A & B columns
        var data = sheet.getDataRange().getValues();
        for (var r = 1; r < data.length; r++) {  // start at r=1 to skip header
          if (String(data[r][0]) === roll && String(data[r][1]) === team) {
            // Write link in column M (13th column)
            sheet.getRange(r + 1, 13).setValue(getThreadLink(thread));
            placed = true;
            break;
          }
        }
      }

      if (!placed) {
        // Couldn't find matching row → log to Errors
        errors.appendRow([roll, team, roundNum, getThreadLink(thread)]);
      }

      // Mark this message as read
      msg.markRead();
    });
  });
}

/**
 * Builds a clickable Gmail link to a thread.
 */
function getThreadLink(thread) {
  var id = thread.getId();
  // This link opens the conversation in the web UI
  return 'https://mail.google.com/mail/u/0/#all/' + id;
}
