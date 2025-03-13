/**
 * Add manual points to a user
 */
function addManualPoints(mnemonic, questionId, points, reason, type) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
    const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);

    if (!scoresSheet || !auditLogSheet) {
        logError('Manual Points', 'Required sheets not found');
        return false;
    }

    try {
        // Generate a unique ID for BONUS entries if it doesn't already have one
        if (questionId === 'BONUS') {
            // Generate a unique timestamp-based ID
            const now = new Date();
            const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
            questionId = `BONUS-${timestamp}`;
        }

        // Get current score and row index
        let currentScore = 0;
        let userRow = -1;
        const scoresData = scoresSheet.getDataRange().getValues();
        
        for (let i = 1; i < scoresData.length; i++) {
            if (scoresData[i][0]?.toLowerCase() === mnemonic.toLowerCase()) {
                currentScore = Number(scoresData[i][3]) || 0;
                userRow = i + 1;  // +1 because array index starts at 0 but sheet rows start at 1
                break;
            }
        }

        if (userRow === -1) {
            throw new Error(`User ${mnemonic} not found in scores sheet`);
        }

        // Calculate new score
        const newScore = currentScore + points;

        // Update score in Scores sheet
        scoresSheet.getRange(userRow, 4).setValue(newScore);  // Column D (4) is the total score

        // Update attempts in Scores sheet
        let attempts = {};
        try {
            const existingAttempts = scoresSheet.getRange(userRow, 6).getValue();
            attempts = JSON.parse(existingAttempts || "{}");
        } catch (e) {
            console.error('Error parsing existing attempts:', e);
        }

        attempts[questionId] = {
            timestamp: new Date(),
            points: points,
            manual: true
        };

        scoresSheet.getRange(userRow, 6).setValue(JSON.stringify(attempts));

        // Log to audit
        const auditEntry = [
            new Date(),           // Timestamp
            mnemonic,            // Mnemonic
            questionId,          // Question ID (now includes unique ID for BONUS)
            reason,              // Answer/Reason
            'Manual',            // Correct?
            'No',               // Duplicate?
            'Yes',              // Correct Role?
            currentScore,        // Previous Points
            points,             // Earned Points
            newScore,           // Total Points
            'Manual Addition'    // Status
        ];
        
        auditLogSheet.appendRow(auditEntry);
        
        // If we have a manual grade log, add it there too
        let processingLogSheet = sheet.getSheetByName("Manual Grade Processing Log");
        if (processingLogSheet) {
            const pointType = questionId.includes("BONUS") ? "Bonus/Recognition Points" : "Question Points";
            
            processingLogSheet.appendRow([
                new Date(),                           // Timestamp
                mnemonic,                            // Mnemonic
                points,                              // Points
                questionId,                          // Question ID (with unique identifier)
                pointType,                           // Point Type
                reason,                              // Reason
                new Date(),                          // Processing Date
                "Direct Addition"                    // Status
            ]);
        }
        
        // Update leaderboard
        updateLeaderboard();
        
        return true;
    } catch (error) {
        logError('Manual Points', `Error adding points for ${mnemonic}: ${error.message}`);
        return false;
    }
}

/**
 * Handle form grading override from Form Responses
 */
function handleFormGradeOverride() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const formResponsesSheet = sheet.getSheetByName(SHEETS.FORM_RESPONSES);
    const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
    
    // Get or create processing log sheet
    let processingLogSheet = sheet.getSheetByName("Manual Grade Processing Log");
    if (!processingLogSheet) {
        processingLogSheet = sheet.insertSheet("Manual Grade Processing Log");
        processingLogSheet.appendRow(["Timestamp", "Mnemonic", "Score", "Row Index", "Processing Date", "Status"]);
        processingLogSheet.setFrozenRows(1);
    }
    
    if (!formResponsesSheet || !scoresSheet || !processingLogSheet) {
        logError('Grade Override', 'Required sheets not found');
        return;
    }

    // Get processed entries from log sheet
    const processedLog = processingLogSheet.getDataRange().getValues();
    const processedRows = new Set(); // Track processed row indices
    
    // Skip header row
    for (let i = 1; i < processedLog.length; i++) {
        if (processedLog[i][3]) { // Row index is stored in column D
            processedRows.add(parseInt(processedLog[i][3]));
        }
    }

    console.log("Processed rows in log:", processedRows.size);

    const responses = formResponsesSheet.getDataRange().getValues();
    const newProcessedEntries = [];
    let processedCount = 0;
    
    // Process each response
    for (let i = 1; i < responses.length; i++) {
        // Skip if this row index has already been processed
        if (processedRows.has(i)) {
            console.log(`Skipping already processed row ${i}`);
            continue;
        }
        
        const row = responses[i];
        const timestamp = row[0];
        const score = row[1];  // Score column B
        const mnemonic = row[2];  // Mnemonic in column C
        
        if (!timestamp || !score || !mnemonic) continue;
        
        // Parse the score fraction (e.g., "4/4" -> 4)
        let points = 0;
        if (typeof score === 'string' && score.includes('/')) {
            const [earned, total] = score.split('/').map(Number);
            points = earned;
        } else if (!isNaN(score)) {
            points = Number(score);
        }

        if (points > 0) {
            console.log(`Processing new grade for ${mnemonic} (row ${i}): ${points} points`);
            
            // Use consistent question ID format for manual grades
            const questionId = 'MANUAL_GRADE';
            
            addManualPoints(
                mnemonic, 
                questionId, 
                points, 
                `Manual grade from form response: ${score} points`
            );
            
            // Add to log of processed entries
            newProcessedEntries.push([
                timestamp,
                mnemonic,
                score,
                i, // Store row index for exact matching
                new Date(),
                "Processed"
            ]);
            
            processedCount++;
        }
    }
    
    // Batch append new processed entries to log
    if (newProcessedEntries.length > 0) {
        processingLogSheet.getRange(
            processingLogSheet.getLastRow() + 1, 
            1, 
            newProcessedEntries.length, 
            newProcessedEntries[0].length
        ).setValues(newProcessedEntries);
        
        console.log(`Successfully processed ${processedCount} new manual grades`);
    } else {
        console.log("No new manual grades to process");
    }
}

/**
 * Show dialog for adding manual points
 */
function showManualPointsDialog() {
  const html = HtmlService.createHtmlOutput(`
    <form id="manualPointsForm">
      <div style="margin-bottom: 10px;">
        <label for="mnemonic">Mnemonic:</label><br>
        <input type="text" id="mnemonic" required style="width: 100%;">
      </div>
      <div style="margin-bottom: 10px;">
        <label for="type">Point Type:</label><br>
        <select id="type" onchange="toggleQuestionId()" style="width: 100%;">
          <option value="question">Question Points</option>
          <option value="bonus">Bonus/Recognition Points</option>
        </select>
      </div>
      <div id="questionIdField" style="margin-bottom: 10px;">
        <label for="questionId">Question ID:</label><br>
        <input type="text" id="questionId" style="width: 100%;">
      </div>
      <div style="margin-bottom: 10px;">
        <label for="points">Points:</label><br>
        <input type="number" id="points" required style="width: 100%;">
      </div>
      <div style="margin-bottom: 10px;">
        <label for="reason">Reason:</label><br>
        <textarea id="reason" required style="width: 100%; height: 60px;" 
          placeholder="Enter reason for points (e.g., 'STAR recognition', 'Extra effort on case study', etc.)"></textarea>
      </div>
      <button onclick="submitForm()" style="width: 100%; padding: 8px;">Add Points</button>
    </form>
    <script>
      function toggleQuestionId() {
        const type = document.getElementById('type').value;
        const questionField = document.getElementById('questionIdField');
        const questionId = document.getElementById('questionId');
        
        if (type === 'bonus') {
          questionField.style.display = 'none';
          questionId.required = false;
          questionId.value = 'BONUS';
        } else {
          questionField.style.display = 'block';
          questionId.required = true;
        }
      }
      
      function submitForm() {
        const form = document.getElementById('manualPointsForm');
        const data = {
          mnemonic: document.getElementById('mnemonic').value,
          questionId: document.getElementById('questionId').value || 'BONUS',
          points: Number(document.getElementById('points').value),
          reason: document.getElementById('reason').value,
          type: document.getElementById('type').value
        };
        
        if (!data.questionId && data.type === 'question') {
          alert('Please enter a Question ID for question points');
          return;
        }
        
        google.script.run
          .withSuccessHandler(() => {
            alert('Points added successfully!');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Error adding points: ' + error);
          })
          .addManualPoints(data.mnemonic, data.questionId, data.points, data.reason);
      }
    </script>
  `)
    .setWidth(400)
    .setHeight(350)
    .setTitle('Add Manual Points');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Manual Points');
}

/**
 * Adds Manual Grade Processing Log 
 */
function setupManualGradeProcessingLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create the processing log sheet
  let processingLogSheet = sheet.getSheetByName("Manual Grade Processing Log");
  if (!processingLogSheet) {
    processingLogSheet = sheet.insertSheet("Manual Grade Processing Log");
    processingLogSheet.appendRow([
      "Timestamp", "Mnemonic", "Points", "Question ID", "Point Type", 
      "Reason", "Processing Date", "Status"
    ]);
    processingLogSheet.setFrozenRows(1);
  }
  
  // Get existing manual grades from audit log
  const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
  if (!auditLogSheet) {
    console.error("Audit Log sheet not found");
    return;
  }
  
  const auditData = auditLogSheet.getDataRange().getValues();
  const entries = [];
  
  // Look for any Manual Addition in the audit log
  for (let i = 1; i < auditData.length; i++) {
    const row = auditData[i];
    if (row[10] === "Manual Addition") {
      // Extract the question ID 
      const questionId = row[2] || "UNKNOWN";
      const pointType = questionId === "BONUS" ? "Bonus/Recognition Points" : "Question Points";
      
      entries.push([
        row[0],                 // Timestamp
        row[1],                 // Mnemonic
        row[8],                 // Points awarded
        questionId,             // Question ID
        pointType,              // Point Type
        row[3],                 // Reason
        new Date(),             // Current date
        "Imported from Audit Log"
      ]);
    }
  }
  
  // Add existing entries to the log
  if (entries.length > 0) {
    processingLogSheet.getRange(
      processingLogSheet.getLastRow() + 1, 
      1, 
      entries.length, 
      entries[0].length
    ).setValues(entries);
    
    console.log(`Imported ${entries.length} manual points entries to log`);
  } else {
    console.log("No manual points found in Audit Log");
  }
}

/**
 * After adding manual points, add to Manual Grade Processing log:
 */
function addToManualGradeLog(timestamp, mnemonic, points, questionId, pointType, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let processingLogSheet = sheet.getSheetByName("Manual Grade Processing Log");
  
  if (!processingLogSheet) {
    setupManualGradeProcessingLog();
    processingLogSheet = sheet.getSheetByName("Manual Grade Processing Log");
  }
  
  processingLogSheet.appendRow([
    timestamp,
    mnemonic,
    points,
    questionId,
    pointType,
    reason,
    new Date(),
    "Direct Addition"
  ]);
}

/**
 * Adds option to recover Bonus Points in case Audit Log is deleted
 */
function recoverBonusPoints() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const processingLogSheet = sheet.getSheetByName("Manual Grade Processing Log");
  
  if (!processingLogSheet) {
    console.error("Manual Grade Processing Log not found");
    return;
  }
  
  const logData = processingLogSheet.getDataRange().getValues();
  const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
  
  if (!scoresSheet) {
    console.error("Scores sheet not found");
    return;
  }
  
  // Get all processed entries from the Manual Grade Processing Log
  const processedEntries = new Set();
  
  // Get processed entries from Audit Log as well 
  const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
  if (auditLogSheet) {
    const auditData = auditLogSheet.getDataRange().getValues();
    // Skip header row
    for (let i = 1; i < auditData.length; i++) {
      const row = auditData[i];
      // Check for both Manual Addition and Recovery Action
      if ((row[10] === "Manual Addition" || row[10] === "Recovery Action") && 
          row[2].includes("BONUS-")) {
        // Create a unique key from mnemonic + bonus ID to check against
        const entryKey = `${row[1]}_${row[2]}`;
        processedEntries.add(entryKey);
      }
    }
  }
  
  // Get bonus entries that haven't been recovered yet
  const bonusEntries = [];
  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    if (row[4] === "Bonus/Recognition Points" && row[7] !== "Recovered") {
      const entryKey = `${row[1]}_${row[3]}`;
      
      // Only process if not already in the audit log
      if (!processedEntries.has(entryKey)) {
        bonusEntries.push({
          row: i + 1,  // Row number in the sheet
          timestamp: row[0],
          mnemonic: row[1],
          points: row[2],
          questionId: row[3],
          reason: row[5]
        });
        processedEntries.add(entryKey);
      } else {
        // Already processed, just update the status
        processingLogSheet.getRange(i + 1, 8).setValue("Recovered");
      }
    }
  }
  
  console.log(`Found ${bonusEntries.length} new bonus entries to process`);
  
  // Apply the bonus points
  let processed = 0;
  for (const entry of bonusEntries) {
    // Instead of using addManualPoints which would add to the log again,
    // directly update the Scores sheet
    const success = directlyUpdateScore(
      entry.mnemonic, 
      entry.questionId, 
      entry.points, 
      entry.reason
    );
    
    if (success) {
      processed++;
      // Mark as recovered in the log
      processingLogSheet.getRange(entry.row, 8).setValue("Recovered");
    }
  }
  
  console.log(`Successfully recovered ${processed} bonus entries`);
  
  // Update leaderboard
  updateLeaderboard();
}

/**
 * Update score directly without adding to logs again
 */
function directlyUpdateScore(mnemonic, questionId, points, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
  const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
  
  if (!scoresSheet) {
    console.error("Scores sheet not found");
    return false;
  }
  
  try {
    // Get current score and row index
    let currentScore = 0;
    let userRow = -1;
    const scoresData = scoresSheet.getDataRange().getValues();
    
    for (let i = 1; i < scoresData.length; i++) {
      if (scoresData[i][0]?.toLowerCase() === mnemonic.toLowerCase()) {
        currentScore = Number(scoresData[i][3]) || 0;
        userRow = i + 1;  // +1 because array index starts at 0 but sheet rows start at 1
        break;
      }
    }

    if (userRow === -1) {
      throw new Error(`User ${mnemonic} not found in scores sheet`);
    }

    // Calculate new score
    const newScore = currentScore + points;

    // Update score in Scores sheet
    scoresSheet.getRange(userRow, 4).setValue(newScore);  // Column D (4) is the total score

    // Update attempts in Scores sheet
    let attempts = {};
    try {
      const existingAttempts = scoresSheet.getRange(userRow, 6).getValue();
      attempts = JSON.parse(existingAttempts || "{}");
    } catch (e) {
      console.error('Error parsing existing attempts:', e);
    }

    // Only add to attempts if this questionId doesn't already exist
    if (!attempts[questionId]) {
      attempts[questionId] = {
        timestamp: new Date(),
        points: points,
        manual: true,
        recovered: true  // Mark as recovered so we know it's a recovery
      };

      scoresSheet.getRange(userRow, 6).setValue(JSON.stringify(attempts));
    }
    
    // Add to audit log with special status
    if (auditLogSheet) {
      const auditEntry = [
        new Date(),           // Timestamp
        mnemonic,            // Mnemonic
        questionId,          // Question ID
        reason || "Recovery of bonus points",  // Reason
        'Manual',            // Correct?
        'No',               // Duplicate?
        'Yes',              // Correct Role?
        currentScore,        // Previous Points
        points,             // Earned Points
        newScore,           // Total Points
        'Recovery Action'    // Special status to differentiate from normal Manual Addition
      ];
      
      auditLogSheet.appendRow(auditEntry);
    }
    
    return true;
  } catch (error) {
    console.error(`Error updating score for ${mnemonic}:`, error);
    return false;
  }
}
