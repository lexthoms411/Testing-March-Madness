/**
 * Main grading function with optimizations
 */
function gradeResponses() {
    // Use lock service to prevent concurrent execution
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) { // 10 seconds timeout
            console.log("Could not obtain lock for grading. Another process is running.");
            return;
        }
        
        console.info("📌 Starting grading process...");
        const startTime = new Date().getTime();
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const responsesSheet = sheet.getSheetByName("Form Responses (Raw)");
        const questionBankSheet = sheet.getSheetByName(SHEETS.QUESTION_BANK);
        const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
        const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);

        if (!responsesSheet || !questionBankSheet || !scoresSheet || !auditLogSheet) {
            console.error("❌ Missing required sheets");
            logError('Grade Responses', 'Missing required sheets');
            return;
        }

        // Get already processed responses with caching
        const processedResponses = getProcessedResponsesWithCache(auditLogSheet);

        // Sync new responses
        syncResponses();

        // Get question data with caching
        const questionMap = getQuestionMapWithCache(questionBankSheet);
        const answerMapping = getAnswerMappingWithCache(questionMap);
        
        // Get valid mnemonics with caching
        const validMnemonics = getValidMnemonicsWithCache(scoresSheet);

        // Get only ungraded responses for efficiency
        const ungradedResponses = getUngradedResponses(responsesSheet, validMnemonics);
        
        if (ungradedResponses.length === 0) {
            console.info("✅ No ungraded responses to process");
            return;
        }
        
        console.info(`🔍 Processing ${ungradedResponses.length} ungraded responses`);

        let auditLogEntries = [];
        const MAX_EXECUTION_TIME = 250000; // 250 seconds (under 300s limit)

        // Process responses
        for (let i = 0; i < ungradedResponses.length; i++) {
            // Check if we're approaching execution time limit
            if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
                console.warn("⚠️ Approaching execution time limit, saving progress");
                
                // Save audit entries processed so far
                if (auditLogEntries.length > 0) {
                    appendToAuditLog(auditLogSheet, auditLogEntries);
                    auditLogEntries = [];
                }
                
                // Schedule continuation
                ScriptApp.newTrigger('continueGrading')
                    .timeBased()
                    .after(1000) // 1 second
                    .create();
                    
                return;
            }
            
            const [rowIndex, row] = ungradedResponses[i];
            const timestamp = row[0];
            const mnemonic = row[1]?.toLowerCase();
            const answerData = parseAnswer(row[2]);
            
            if (!mnemonic || !answerData) continue;

            for (const [qID, userAnswer] of Object.entries(answerData)) {
                const responseKey = `${timestamp}_${mnemonic}_${qID}`.toLowerCase();

                if (processedResponses.has(responseKey)) {
                    continue;
                }

                const questionData = questionMap[qID];
                if (!questionData) {
                    console.warn(`Question ${qID} not found in bank`);
                    logError('Grade Response', `Question ${qID} not found in bank`);
                    continue;
                }

                // Get current score before grading
                const currentScore = getCurrentScore(scoresSheet, mnemonic);

                // Get actual role from the Scores sheet
                const actualRole = getUserRole(scoresSheet, mnemonic).trim().toLowerCase();
                const requiredRole = (questionData.targetRole || "").trim().toLowerCase();

                // Check both role mismatch & duplicate attempt at the same time
                const correctRole = actualRole === requiredRole || !requiredRole;
                const isDuplicate = hasAttemptedBefore(scoresSheet, mnemonic, qID);

                // Grade answer regardless of eligibility
                const isCorrect = isAnswerCorrect(userAnswer, questionData.correctAnswer, questionData.type);
                let earnedPoints = 0;

                // Only award points if eligible (correct role and not duplicate)
                if (correctRole && !isDuplicate) {
                    if (questionData.type && questionData.type.toLowerCase() === "multiple select") {
                        earnedPoints = calculatePartialCredit(
                            userAnswer,
                            questionData.correctAnswer,
                            questionData.type,
                            questionData.points
                        );
                    } else {
                        earnedPoints = isCorrect ? questionData.points : 0;
                    }

                    // Update scores
                    updateScores(scoresSheet, mnemonic, qID, earnedPoints, timestamp);
                }

                // Update raw responses with correct/incorrect status
                responsesSheet.getRange(rowIndex + 1, 6).setValue(isCorrect ? "Correct" : "Incorrect");

                // Get shortened answers for display
                let formattedUserAnswer = getAnswerLetters(userAnswer, qID, answerMapping);
                let formattedCorrectAnswer = getAnswerLetters(questionData.correctAnswer, qID, answerMapping);
                
                // Short display version for audit log
                const answerDisplay = `Answer: ${shortenAnswerText(formattedUserAnswer)} (Expected: ${shortenAnswerText(formattedCorrectAnswer)})`;

                // Log to audit with new column structure
                auditLogEntries.push([
                    timestamp,                    // Timestamp
                    mnemonic,                    // Mnemonic
                    qID,                         // Question ID
                    answerDisplay,               // Shortened answer
                    isCorrect ? "Correct" : "Incorrect",  // Correct? (now shows regardless of status)
                    isDuplicate ? "Yes" : "No",  // Duplicate Attempt?
                    correctRole ? "Yes" : "No",  // Correct Role?
                    currentScore,                // Previous Points
                    earnedPoints,                // Earned Points
                    currentScore + earnedPoints, // Total Points
                    isDuplicate ? "Duplicate" :
                        !correctRole ? "Role Mismatch" :
                        "Processed"              // Status
                ]);
                
                // Process in smaller batches to avoid memory issues
                if (auditLogEntries.length >= 50) {
                    appendToAuditLog(auditLogSheet, auditLogEntries);
                    auditLogEntries = [];
                }
            }

            // Mark as graded
            responsesSheet.getRange(rowIndex + 1, 5).setValue("Yes");
        }

        // Add any remaining audit entries
        if (auditLogEntries.length > 0) {
            appendToAuditLog(auditLogSheet, auditLogEntries);
        }

        // Update audit log formatting
        updateAuditLogFormatting();
        
        // Update leaderboards
        updateLeaderboard();
        
        // Update Processed Responses
        updateProcessedResponses();

        console.info("🎉 Grading complete!");
    } catch (e) {
        console.error("❌ Error in grading process:", e.message, e.stack);
        logError('Grade Responses', `Error in grading process: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Continue grading from where we left off
 */
function continueGrading() {
    // Delete all triggers for this function to prevent duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'continueGrading') {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
    
    // Continue grading
    gradeResponses();
}

/**
 * Get ungraded responses efficiently
 */
function getUngradedResponses(responsesSheet, validMnemonics) {
    const data = responsesSheet.getDataRange().getValues();
    const result = [];
    
    // Create a Set for faster lookups
    const validMnemonicsSet = new Set(validMnemonics);
    
    for (let i = 1; i < data.length; i++) {
        const mnemonic = data[i][1]?.toLowerCase();
        if (data[i][4] !== "Yes" && mnemonic && validMnemonicsSet.has(mnemonic)) {
            result.push([i, data[i]]);
        }
    }
    
    return result;
}

/**
 * Append to audit log with retry
 */
function appendToAuditLog(auditLogSheet, entries) {
    if (!entries || entries.length === 0) return;
    
    return retryOperation(() => {
        const lastRow = auditLogSheet.getLastRow();
        auditLogSheet.getRange(
            lastRow + 1, 
            1, 
            entries.length, 
            entries[0].length
        ).setValues(entries);
    });
}

/**
 * Get processed responses with caching
 */
function getProcessedResponsesWithCache(auditLogSheet) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'processedResponses';
    
    // Try to get from cache first
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
        try {
            return new Set(JSON.parse(cachedData));
        } catch (e) {
            console.warn("⚠️ Cache parse error, rebuilding processed responses");
        }
    }
    
    // If not in cache or parse error, rebuild
    const processedResponses = new Set();
    const auditData = auditLogSheet.getDataRange().getValues();

    for (let i = 1; i < auditData.length; i++) {
        const key = `${auditData[i][0]}_${auditData[i][1]}_${auditData[i][2]}`.toLowerCase();
        processedResponses.add(key);
    }
    
    // Store in cache for 6 hours (needs to be serialized as array)
    // Since there are size limitations, we'll only cache if it's not too large
    if (processedResponses.size < 5000) {
        cache.put(cacheKey, JSON.stringify(Array.from(processedResponses)), 21600);
    }

    return processedResponses;
}

/**
 * Get question map with caching
 */
function getQuestionMapWithCache(questionBankSheet) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'questionMap';
    
    // Try to get from cache first
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
        try {
            return JSON.parse(cachedData);
        } catch (e) {
            console.warn("⚠️ Cache parse error, rebuilding question map");
        }
    }
    
    // If not in cache or parse error, rebuild
    const questionBankData = questionBankSheet.getDataRange().getValues();
    const questionMap = {};
    
    for (let i = 1; i < questionBankData.length; i++) {
        const row = questionBankData[i];
        const qID = row[1];
        if (qID) {
            questionMap[qID] = {
                question: row[2],
                correctAnswer: row[9],
                type: row[10],
                targetRole: row[11],
                points: parseInt(row[12]) || 0
            };
        }
    }
    
    // Cache for 6 hours
    cache.put(cacheKey, JSON.stringify(questionMap), 21600);
    
    return questionMap;
}

/**
 * Get answer mapping with caching
 */
function getAnswerMappingWithCache(questionMap) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'answerMapping';
    
    // Try to get from cache first
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
        try {
            return JSON.parse(cachedData);
        } catch (e) {
            console.warn("⚠️ Cache parse error, rebuilding answer mapping");
        }
    }
    
    // If not in cache or parse error, rebuild
    const answerMapping = {};
    
    // Need to get question bank data for options
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const questionBankSheet = sheet.getSheetByName(SHEETS.QUESTION_BANK);
    const questionBankData = questionBankSheet.getDataRange().getValues();
    
    for (let i = 1; i < questionBankData.length; i++) {
        const row = questionBankData[i];
        const qID = row[1];
        const type = row[10];
        
        if (qID && type && type.toLowerCase().includes("multiple")) {
            const options = [row[3], row[4], row[5], row[6], row[7], row[8]].filter(Boolean);
            const letterMap = {};
            
            options.forEach((text, index) => {
                if (text) {
                    const letter = String.fromCharCode(65 + index); // A, B, C, etc.
                    letterMap[text.toLowerCase().trim()] = letter;
                }
            });
            
            if (Object.keys(letterMap).length > 0) {
                answerMapping[qID] = letterMap;
            }
        }
    }
    
    // Cache for 6 hours
    cache.put(cacheKey, JSON.stringify(answerMapping), 21600);
    
    return answerMapping;
}

/**
 * Get valid mnemonics with caching
 */
function getValidMnemonicsWithCache(scoresSheet) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'validMnemonics';
    
    // Try to get from cache first
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
        try {
            return JSON.parse(cachedData);
        } catch (e) {
            console.warn("⚠️ Cache parse error, rebuilding valid mnemonics");
        }
    }
    
    // If not in cache or parse error, rebuild
    const validMnemonics = scoresSheet.getRange('A2:A')
        .getValues()
        .map(row => row[0]?.toLowerCase())
        .filter(Boolean);
    
    // Cache for 6 hours
    cache.put(cacheKey, JSON.stringify(validMnemonics), 21600);
    
    return validMnemonics;
}

/**
 * Sync form responses to raw responses sheet with improved error handling
 */
function syncResponses() {
    // Use lock service to prevent concurrent execution
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) { // 10 seconds timeout
            console.log("Could not obtain lock for syncing. Another process is running.");
            return [];
        }
        
        console.info("🔄 Starting response sync...");

        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const formResponsesSheet = sheet.getSheetByName(SHEETS.FORM_RESPONSES);
        const rawResponsesSheet = sheet.getSheetByName("Form Responses (Raw)");

        if (!formResponsesSheet || !rawResponsesSheet) {
            console.error("❌ Required sheets not found");
            logError('Sync Responses', 'Required sheets not found');
            return [];
        }

        // Get last sync timestamp
        const props = PropertiesService.getScriptProperties();
        const lastSyncKey = 'lastSyncTimestamp';
        let lastTimestamp = null;
        
        // Check if raw sheet is empty
        const rawRowCount = rawResponsesSheet.getLastRow();
        const isEmpty = rawRowCount <= 1;
        
        if (!isEmpty) {
            const lastSyncTimestamp = props.getProperty(lastSyncKey);
            if (lastSyncTimestamp) {
                lastTimestamp = new Date(lastSyncTimestamp);
            }
        }

        // Get form data headers
        const headers = formResponsesSheet.getRange(1, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];
        
        // Get existing entries to avoid duplicates
        const rawData = rawResponsesSheet.getDataRange().getValues();
        const existingEntries = new Set();
        
        if (!isEmpty) {
            for (let i = 1; i < rawData.length; i++) {
                if (rawData[i][0] && rawData[i][1] && rawData[i][2]) {
                    const key = `${rawData[i][0]}_${String(rawData[i][1]).toLowerCase()}_${rawData[i][2]}`;
                    existingEntries.add(key);
                }
            }
        }

        // Get form data (all or just new based on timestamp)
        const formData = formResponsesSheet.getDataRange().getValues();
        const newResponses = [];
        let newTimestamp = lastTimestamp || new Date(0);
        let skippedRows = 0;
        let invalidMnemonics = 0;

        // Process each form response
        for (let i = 1; i < formData.length; i++) {
            const row = formData[i];
            const timestamp = row[0];
            
            // Skip if older than last sync time
            if (lastTimestamp && timestamp && timestamp < lastTimestamp) {
                skippedRows++;
                continue;
            }
            
            // Track newest timestamp
            if (timestamp && timestamp > newTimestamp) {
                newTimestamp = timestamp;
            }

            // Look for mnemonic in the correct column (column C, index 2)
            const mnemonic = row[2]; 
            
            // Better validation of mnemonic
            if (!mnemonic || typeof mnemonic !== 'string' || mnemonic.trim() === '') {
                console.warn(`⚠️ Invalid mnemonic at row ${i + 1}`);
                invalidMnemonics++;
                continue;
            }

            const mnemonicLower = mnemonic.toString().toLowerCase().trim();
            if (mnemonicLower === '') {
                console.warn(`⚠️ Empty mnemonic after trimming at row ${i + 1}`);
                invalidMnemonics++;
                continue;
            }
            
            // Get role from column D (index 3)
            const role = row[3] || '';

            // Process answers - collect question answers from remaining columns
            let answerDataObj = {};
            for (let col = 4; col < headers.length; col++) {
                if (!headers[col]) continue; // Skip columns with no header
                
                const answer = row[col];
                if (answer && answer.toString().trim() !== "") {
                    const questionID = extractQuestionID(headers[col]);
                    if (questionID) {
                        answerDataObj[questionID] = answer.toString().trim();
                    }
                }
            }

            // Only add if we have at least one answer
            if (Object.keys(answerDataObj).length > 0) {
                const answerJson = JSON.stringify(answerDataObj);
                const entryKey = `${timestamp}_${mnemonicLower}_${answerJson}`;

                // If raw sheet is empty OR this entry doesn't exist yet, add it
                if (isEmpty || !existingEntries.has(entryKey)) {
                    const formattedRow = [timestamp, mnemonicLower, answerJson, role, "No", ""];
                    newResponses.push(formattedRow);
                    existingEntries.add(entryKey); // Add to set to avoid duplicates in this run
                }
            }
        }

        // Add new responses in batches
        if (newResponses.length > 0) {
            const BATCH_SIZE = 50;
            for (let i = 0; i < newResponses.length; i += BATCH_SIZE) {
                const batch = newResponses.slice(i, Math.min(i + BATCH_SIZE, newResponses.length));
                rawResponsesSheet.getRange(
                    rawResponsesSheet.getLastRow() + 1, 
                    1, 
                    batch.length, 
                    batch[0].length
                ).setValues(batch);
            }
            
            // Store newest timestamp for next sync
            if (newTimestamp > new Date(0)) {
                props.setProperty(lastSyncKey, newTimestamp.toISOString());
            }
            
            // Apply consistent formatting to new entries
            fixTimestampFormatting();
            fixTextAlignment();
        }

        console.info(`✅ Synced ${newResponses.length} new responses. Skipped ${skippedRows} old rows and ${invalidMnemonics} invalid mnemonics.`);
        return newResponses;
    } catch (e) {
        console.error("❌ Error in syncResponses:", e.message, e.stack);
        logError('Sync Responses', `Error syncing: ${e.message}\n${e.stack}`);
        return [];
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Update scores in scores sheet with retry logic
 */
function updateScores(scoresSheet, mnemonic, questionID, points, timestamp) {
    return retryOperation(() => {
        // Ensure scoresSheet is properly defined
        if (!scoresSheet) {
            console.error("❌ scoresSheet is undefined. Attempting to retrieve it.");
            const sheet = SpreadsheetApp.getActiveSpreadsheet();
            scoresSheet = sheet.getSheetByName(SHEETS.SCORES);

            if (!scoresSheet) {
                console.error("❌ Scores sheet not found!");
                logError('Update Scores', 'Scores sheet not found');
                return;
            }
        }

        const scoresData = scoresSheet.getDataRange().getValues();
        const mnemonicLower = mnemonic.toLowerCase();

        for (let i = 1; i < scoresData.length; i++) {
            const row = scoresData[i];
            if (!row || row.length < 4 || !row[0]) continue; // Skip empty or malformed rows

            if (row[0].toLowerCase() === mnemonicLower) {
                if (!hasAttemptedBefore(scoresSheet, mnemonic, questionID)) {
                    let currentScore = row[3] || 0;
                    let newScore = currentScore + points;
                    scoresSheet.getRange(i + 1, 4).setValue(newScore);

                    let attempts = {};
                    try {
                        attempts = JSON.parse(row[5] || "{}");
                    } catch (e) {
                        console.error(`❌ Error parsing attempts for ${mnemonic}:`, e);
                        logError('Update Scores', `Error parsing attempts for ${mnemonic}: ${e.message}`);
                    }

                    attempts[questionID] = { timestamp, points };
                    scoresSheet.getRange(i + 1, 6).setValue(JSON.stringify(attempts));

                    console.info(`✅ Updated score for ${mnemonic}: ${newScore} (Question: ${questionID}, Points: ${points})`);
                } else {
                    console.info(`ℹ️ Skipped score update - ${mnemonic} already attempted question ${questionID}`);
                }
                return;
            }
        }
    });
}

/**
 * Optimized Reset all scores to zero with retry logic
 */
function resetAllScores() {
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) {
            console.log("Could not obtain lock for resetting scores. Another process is running.");
            return;
        }
        
        console.info("🔄 Resetting all scores...");

        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);

        if (!scoresSheet) {
            console.error("❌ Scores sheet not found");
            logError('Reset Scores', 'Scores sheet not found');
            return;
        }

        const lastRow = scoresSheet.getLastRow();
        if (lastRow <= 1) {
            console.info("ℹ️ No scores to reset.");
            return;
        }

        // Batch update using setValues for efficiency
        const numRows = lastRow - 1;
        const zeroArray = Array(numRows).fill([0]);
        const emptyJsonArray = Array(numRows).fill(["{}"]);

        retryOperation(() => {
            scoresSheet.getRange(2, 4, numRows, 1).setValues(zeroArray); // Reset scores
            scoresSheet.getRange(2, 6, numRows, 1).setValues(emptyJsonArray); // Reset attempts
        });

        // Update leaderboards
        updateLeaderboard();

        // Clear caches
        clearCaches();

        console.info(`✅ Successfully reset scores for ${numRows} participants.`);
    } catch (e) {
        console.error("❌ Error in resetAllScores:", e.message, e.stack);
        logError('Reset Scores', `Error resetting scores: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Clear script caches
 */
function clearCaches() {
    const cache = CacheService.getScriptCache();
    const keysToDelete = [
        'processedResponses',
        'questionMap',
        'answerMapping',
        'validMnemonics'
    ];
    
    for (const key of keysToDelete) {
        cache.remove(key);
    }
    
    console.log("🧹 Cleared all caches");
}

/**
 * Optimized Delete all synced response data with proper locking
 */
function deleteSyncData() {
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) {
            console.log("Could not obtain lock for deleting sync data. Another process is running.");
            return;
        }
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const rawResponsesSheet = sheet.getSheetByName("Form Responses (Raw)");

        if (!rawResponsesSheet) {
            logError('Delete Sync Data', 'Raw responses sheet not found');
            return;
        }

        const lastRow = rawResponsesSheet.getLastRow();
        const lastColumn = rawResponsesSheet.getLastColumn();

        if (lastRow > 1 && lastColumn > 0) {
            // Clear all data except header with retry
            retryOperation(() => {
                rawResponsesSheet.getRange(2, 1, lastRow - 1, lastColumn).clear();
            });
            
            // Also clear the last sync timestamp to force full resync
            PropertiesService.getScriptProperties().deleteProperty('lastSyncTimestamp');
            
            // Clear caches
            clearCaches();
            
            console.info(`✅ Successfully deleted ${lastRow - 1} synced responses.`);
        } else {
            console.info("ℹ️ No synced responses to delete.");
        }
    } catch (e) {
        console.error("❌ Error in deleteSyncData:", e.message, e.stack);
        logError('Delete Sync Data', `Error deleting sync data: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Clear audit log for testing purposes with proper locking
 */
function clearAuditLog() {
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) {
            console.log("Could not obtain lock for clearing audit log. Another process is running.");
            return;
        }
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);

        if (!auditLogSheet) {
            console.error("❌ Audit log sheet not found");
            return;
        }

        const lastRow = auditLogSheet.getLastRow();
        if (lastRow > 1) {
            // Clear all data except header with retry
            retryOperation(() => {
                auditLogSheet.getRange(2, 1, lastRow - 1, auditLogSheet.getLastColumn()).clear();
            });
            
            // Clear cache
            CacheService.getScriptCache().remove('processedResponses');
            
            console.info(`✅ Successfully cleared ${lastRow - 1} audit log entries.`);
        } else {
            console.info("ℹ️ No audit log entries to clear.");
        }
    } catch (e) {
        console.error("❌ Error in clearAuditLog:", e.message, e.stack);
        logError('Clear Audit Log', `Error clearing audit log: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Setup all required triggers for the competition
 */
function setupTriggers() {
    // Clear existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }

    
    
    // Create time trigger to process the queue every 5 minutes
    ScriptApp.newTrigger('processQueue')
        .timeBased()
        .everyMinutes(5)
        .create();
    
    // Get the form ID from the constants
    try {
        const form = FormApp.openById(FORM_ID);
        ScriptApp.newTrigger('onFormSubmit')
            .forForm(form)
            .onFormSubmit()
            .create();
        console.log("✅ Form trigger created successfully");
    } catch (e) {
        console.error("❌ Error creating form trigger: " + e.message);
    }
    
    // Create hourly trigger to archive old data
    ScriptApp.newTrigger('archiveOldData')
        .timeBased()
        .everyHours(12)
        .create();

    ScriptApp.newTrigger('updateDailyQuestions')
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .create();
        
    console.log("✅ All triggers set up successfully");
}

/**
 * Handle form submission - just add to queue, don't process yet
 */
function onFormSubmit(e) {
    try {
        // Add debug info
        console.log("Form submission received");
        
        // Different handling based on whether this is a form trigger or manual run
        let timestamp = new Date();
        let mnemonic = "";
        
        if (e && e.namedValues) {
            // This is a proper form trigger
            console.log("Processing form trigger");
            
            // Get mnemonic - assuming the field is labeled "Mnemonic" in the form
            const mnemonicArray = e.namedValues['Mnemonic'] || [];
            mnemonic = mnemonicArray.length > 0 ? mnemonicArray[0].toLowerCase().trim() : "";
            
            // Get timestamp
            timestamp = e.values ? new Date(e.values[0]) : timestamp;
        } else if (e && e.response) {
            // This is a form submission object
            console.log("Processing form response object");
            
            // Get submission data
            timestamp = e.response.getTimestamp();
            const itemResponses = e.response.getItemResponses();
            
            // Find the mnemonic - assumes your form has a specific question for mnemonic
            for (let i = 0; i < itemResponses.length; i++) {
                const item = itemResponses[i];
                const title = item.getItem().getTitle();
                
                if (title.toLowerCase().includes("mnemonic")) {
                    mnemonic = item.getResponse().toLowerCase().trim();
                    break;
                }
            }
        } else {
            // No event data
            console.warn("No form data available");
            return;
        }
        
        if (!mnemonic) {
            console.warn("⚠️ No mnemonic found in submission");
            return;
        }
        
        // Add to processing queue with retry
        retryOperation(() => {
            const sheet = SpreadsheetApp.getActiveSpreadsheet();
            const queueSheet = sheet.getSheetByName("Processing Queue") || 
                            sheet.insertSheet("Processing Queue");
            
            // If queue sheet is new, add headers
            if (queueSheet.getLastRow() === 0) {
                queueSheet.appendRow(["Timestamp", "Mnemonic", "Processed", "Processing Timestamp"]);
            }
            
            // Add to queue
            queueSheet.appendRow([timestamp, mnemonic, "No", ""]);
        });
        
        console.log(`✅ Added submission from ${mnemonic} to processing queue`);
    } catch (e) {
        console.error("❌ Error in form submission handler:", e.message, e.stack);
        logError('Form Submit', `Error handling submission: ${e.message}`);
    }
}

/**
 * Process the queue of submissions (runs every 5 minutes)
 */
function processQueue() {
    // Use lock service to prevent concurrent execution
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) { // 10 seconds timeout
            console.log("Could not obtain lock for processing queue. Another process is running.");
            return;
        }
        
        console.log("🔄 Processing submission queue...");
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const queueSheet = sheet.getSheetByName("Processing Queue");
        
        if (!queueSheet) {
            console.log("ℹ️ No queue sheet found - nothing to process");
            return;
        }
        
        const queueData = queueSheet.getDataRange().getValues();
        if (queueData.length <= 1) {
            console.log("ℹ️ No items in queue to process");
            return;
        }
        
        // Count how many we need to process
        let pendingCount = 0;
        const pendingIndices = [];
        
        for (let i = 1; i < queueData.length; i++) {
            if (queueData[i][2] === "No") {
                pendingCount++;
                pendingIndices.push(i);
            }
        }
        
        if (pendingCount === 0) {
            console.log("ℹ️ No pending items in queue");
            return;
        }
        
        console.log(`🔄 Found ${pendingCount} queued submissions to process...`);
        
        // Track execution time
        const startTime = new Date().getTime();
        const MAX_EXECUTION_TIME = 250000; // 250 seconds (under 300s limit)
        
        // Process in batches to avoid timeout
        const BATCH_SIZE = 10;
        const batchesToProcess = Math.min(BATCH_SIZE, pendingCount);
        const batchIndices = pendingIndices.slice(0, batchesToProcess);
        
        console.log(`Processing first ${batchesToProcess} of ${pendingCount} pending items`);
        
        // First sync all responses to make sure we have the latest data
        syncResponses();
        
        // Process only a batch of submissions
        processBatchFromQueue(batchIndices, queueSheet, queueData);
        
        // Check if we've used too much time
        const timeElapsed = new Date().getTime() - startTime;
        
        // If more remain and we have time, set up a trigger to continue processing
        if (pendingCount > BATCH_SIZE) {
            if (timeElapsed > MAX_EXECUTION_TIME) {
                console.log(`⏱️ Time limit approaching. Will process remaining ${pendingCount - BATCH_SIZE} items in next cycle`);
            } else {
                console.log(`${pendingCount - BATCH_SIZE} items remain in queue for next processing cycle`);
                
                // Optionally trigger another immediate run for large backlogs
                if (pendingCount > BATCH_SIZE * 3) {
                    ScriptApp.newTrigger('processQueue')
                        .timeBased()
                        .after(60000) // 1 minute
                        .create();
                    console.log("⏱️ Scheduled additional queue processing in 1 minute due to large backlog");
                }
            }
        }
        
        // Update processed responses sheet
        updateProcessedResponses();
    } catch (e) {
        console.error("❌ Error in processQueue:", e.message, e.stack);
        logError('Process Queue', `Error processing queue: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Process a batch of submissions from the queue with retry
 */
function processBatchFromQueue(indices, queueSheet, queueData) {
    // Ensure indices is an array
    if (!Array.isArray(indices)) {
        console.error("❌ indices is not an array:", indices);
        return;
    }
    
    console.log("Processing batch with indices:", indices);

    // Mark these as being processed
    const now = new Date();
    const pendingRows = [];
    const pendingRowIndices = [];
    
    for (let i = 0; i < indices.length; i++) {
        pendingRows.push(["Yes", now]);
        pendingRowIndices.push(indices[i]+1);
    }
    
    // Batch update status
    if (pendingRows.length > 0) {
        try {
            // Update each row one by one to avoid range errors
            for (let i = 0; i < pendingRows.length; i++) {
                retryOperation(() => {
                    queueSheet.getRange(pendingRowIndices[i], 3, 1, 2).setValues([pendingRows[i]]);
                });
            }
            
            // Process these mnemonics
            for (let i = 0; i < indices.length; i++) {
                const index = indices[i];
                // Check if queueData[index] exists and has a second element
                if (queueData[index] && queueData[index][1]) {
                    const mnemonic = queueData[index][1];
                    gradeResponsesForMnemonic(mnemonic);
                } else {
                    console.error(`❌ Invalid queue data at index ${index}`);
                }
            }
        } catch (e) {
            console.error("❌ Error in processBatchFromQueue:", e.message, e.stack);
            logError('Process Batch', `Error processing batch: ${e.message}\n${e.stack}`);
        }
    }
    
    console.log(`✅ Processed ${pendingRows.length} queued submissions`);
}

/**
 * Grade responses for a specific mnemonic with retry
 */
function gradeResponsesForMnemonic(mnemonic) {
    // Check if mnemonic is valid
    if (!mnemonic) {
        console.error("❌ Invalid mnemonic provided to gradeResponsesForMnemonic");
        return;
    }
    
    console.log(`Processing submissions for ${mnemonic}`);
    
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const responsesSheet = sheet.getSheetByName("Form Responses (Raw)");
        
        if (!responsesSheet) {
            console.error("❌ Missing Form Responses (Raw) sheet");
            return;
        }
        
        // Find all ungraded responses for this mnemonic
        const data = responsesSheet.getDataRange().getValues();
        let rowsToGrade = [];
        
        for (let i = 1; i < data.length; i++) {
            // Ensure data[i][1] exists before calling toLowerCase()
            const rowMnemonic = data[i][1] ? data[i][1].toString().toLowerCase() : "";
            
            if (rowMnemonic === mnemonic.toString().toLowerCase() && data[i][4] !== "Yes") {
                rowsToGrade.push(i);
            }
        }
        
        if (rowsToGrade.length === 0) {
            console.log(`No ungraded responses found for ${mnemonic}`);
            return;
        }
        
        console.log(`Found ${rowsToGrade.length} responses to grade for ${mnemonic}`);
        
        // Grade the responses directly
        const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
        const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
        const questionBankSheet = sheet.getSheetByName(SHEETS.QUESTION_BANK);
        
        if (!auditLogSheet || !scoresSheet || !questionBankSheet) {
            console.error("❌ Missing required sheets for grading");
            return;
        }
        
        // Get processed responses
        const processedResponses = getProcessedResponsesWithCache(auditLogSheet);
        
        // Get question data
        const questionMap = getQuestionMapWithCache(questionBankSheet);
        const answerMapping = getAnswerMappingWithCache(questionMap);
        
        // Process each ungraded response for this mnemonic
        let auditLogEntries = [];
        
        for (const rowIndex of rowsToGrade) {
            const row = data[rowIndex];
            const timestamp = row[0];
            const answerData = parseAnswer(row[2]);
            
            if (!answerData) continue;
            
            for (const [qID, userAnswer] of Object.entries(answerData)) {
                const responseKey = `${timestamp}_${mnemonic}_${qID}`.toLowerCase();
                
                if (processedResponses.has(responseKey)) {
                    continue;
                }
                
                const questionData = questionMap[qID];
                if (!questionData) {
                    console.warn(`Question ${qID} not found in bank`);
                    continue;
                }
                
                // Get current score
                const currentScore = getCurrentScore(scoresSheet, mnemonic);
                
                // Check role and duplicate attempt
                const actualRole = getUserRole(scoresSheet, mnemonic).trim().toLowerCase();
                const requiredRole = (questionData.targetRole || "").trim().toLowerCase();
                const correctRole = actualRole === requiredRole || !requiredRole;
                const isDuplicate = hasAttemptedBefore(scoresSheet, mnemonic, qID);
                
                // Grade answer
                const isCorrect = isAnswerCorrect(userAnswer, questionData.correctAnswer, questionData.type);
                let earnedPoints = 0;
                
                // Award points if eligible
                if (correctRole && !isDuplicate) {
                    if (questionData.type && questionData.type.toLowerCase() === "multiple select") {
                        earnedPoints = calculatePartialCredit(
                            userAnswer,
                            questionData.correctAnswer,
                            questionData.type,
                            questionData.points
                        );
                    } else {
                        earnedPoints = isCorrect ? questionData.points : 0;
                    }
                    
                    // Update scores
                    updateScores(scoresSheet, mnemonic, qID, earnedPoints, timestamp);
                }
                
                // Mark as graded in raw responses
                retryOperation(() => {
                    responsesSheet.getRange(rowIndex + 1, 6).setValue(isCorrect ? "Correct" : "Incorrect");
                });
                
                // Get shortened answers for display
                let formattedUserAnswer = getAnswerLetters(userAnswer, qID, answerMapping);
                let formattedCorrectAnswer = getAnswerLetters(questionData.correctAnswer, qID, answerMapping);
                
                // Short display version for audit log
                const answerDisplay = `Answer: ${shortenAnswerText(formattedUserAnswer)} (Expected: ${shortenAnswerText(formattedCorrectAnswer)})`;
                
                // Add to audit log entries
                auditLogEntries.push([
                    timestamp,                    // Timestamp
                    mnemonic,                    // Mnemonic
                    qID,                         // Question ID
                    answerDisplay,               // Shortened answer
                    isCorrect ? "Correct" : "Incorrect",  // Correct?
                    isDuplicate ? "Yes" : "No",  // Duplicate?
                    correctRole ? "Yes" : "No",  // Correct Role?
                    currentScore,                // Previous Points
                    earnedPoints,                // Earned Points
                    currentScore + earnedPoints, // Total Points
                    isDuplicate ? "Duplicate" :
                        !correctRole ? "Role Mismatch" :
                        "Processed"              // Status
                ]);
                
                // Process in smaller batches
                if (auditLogEntries.length >= 20) {
                    appendToAuditLog(auditLogSheet, auditLogEntries);
                    auditLogEntries = [];
                }
            }
            
            // Mark as graded
            retryOperation(() => {
                responsesSheet.getRange(rowIndex + 1, 5).setValue("Yes");
            });
        }
        
        // Add remaining audit entries
        if (auditLogEntries.length > 0) {
            appendToAuditLog(auditLogSheet, auditLogEntries);
        }
        
        // Update audit log formatting
        updateAuditLogFormatting();
        
    } catch (e) {
        console.error(`❌ Error processing ${mnemonic}:`, e.message, e.stack);
        logError('Grade For Mnemonic', `Error processing ${mnemonic}: ${e.message}\n${e.stack}`);
    }
}

/**
 * Update the Processed Responses sheet with incremental updates
 */
function updateProcessedResponses() {
    console.log("🔄 Updating Processed Responses sheet...");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const rawResponsesSheet = sheet.getSheetByName("Form Responses (Raw)");
    const processedSheet = sheet.getSheetByName("Processed Responses") || 
                          sheet.insertSheet("Processed Responses");
    
    if (!rawResponsesSheet) {
        console.error("❌ Form Responses (Raw) sheet not found");
        return;
    }
    
    // Setup headers if needed
    if (processedSheet.getLastRow() === 0) {
        processedSheet.appendRow(["Timestamp", "Mnemonic", "Processed", "Processing Timestamp"]);
    }
    
    // Get existing processed entries to avoid duplicates
    const existingData = processedSheet.getDataRange().getValues();
    const existingEntries = new Set();
    
    for (let i = 1; i < existingData.length; i++) {
        const key = `${existingData[i][0]}_${String(existingData[i][1]).toLowerCase()}`;
        existingEntries.add(key);
    }
    
    // Get all graded responses
    const rawData = rawResponsesSheet.getDataRange().getValues();
    const newProcessedResponses = [];
    
    for (let i = 1; i < rawData.length; i++) {
        if (rawData[i][4] === "Yes") { // If graded
            const key = `${rawData[i][0]}_${String(rawData[i][1]).toLowerCase()}`;
            
            // Only add if not already in processed sheet
            if (!existingEntries.has(key)) {
                newProcessedResponses.push([
                    rawData[i][0], // Timestamp
                    rawData[i][1], // Mnemonic
                    "Yes",         // Processed
                    new Date()     // Processing Timestamp
                ]);
                
                // Add to set to prevent duplicates in this run
                existingEntries.add(key);
            }
        }
    }
    
    // Batch append new entries to processed sheet
    if (newProcessedResponses.length > 0) {
        retryOperation(() => {
            processedSheet.getRange(
                processedSheet.getLastRow() + 1, 
                1, 
                newProcessedResponses.length, 
                4
            ).setValues(newProcessedResponses);
        });
        
        console.log(`✅ Added ${newProcessedResponses.length} new processed responses`);
        
        // Apply consistent formatting to new entries
        fixTimestampFormatting();
        fixTextAlignment();
    } else {
        console.log("ℹ️ No new processed responses to add");
    }
}

/**
 * Update formatting in audit log for better readability
 */
function updateAuditLogFormatting() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    
    if (!auditSheet) return;
    
    // Only update formatting periodically to avoid excess API calls
    const props = PropertiesService.getScriptProperties();
    const lastFormatTime = props.getProperty('lastAuditFormatTime');
    const now = new Date().getTime();
    
    if (lastFormatTime && now - parseInt(lastFormatTime) < 300000) { // 5 minutes
        return; // Skip if formatted recently
    }
    
    // Clear existing rules
    auditSheet.clearConditionalFormatRules();
    
    // Get last row
    const lastRow = Math.max(auditSheet.getLastRow(), 1);
    
    // Create rules array
    const rules = [];
    
    // Status column (column K or 11)
    const statusColumn = 11;
    const statusRange = auditSheet.getRange(2, statusColumn, lastRow - 1, 1);
    
    // Duplicate attempts - yellow
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Duplicate")
        .setBackground("#FFF2CC")
        .setRanges([statusRange])
        .build());
    
    // Role mismatch - orange
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Role Mismatch")
        .setBackground("#FCE5CD")
        .setRanges([statusRange])
        .build());
    
    // Processed - green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Processed")
        .setBackground("#D9EAD3")
        .setRanges([statusRange])
        .build());
    
    // Manual addition - blue
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Manual")
        .setBackground("#CFE2F3")
        .setRanges([statusRange])
        .build());
    
    // Errors - red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Error")
        .setBackground("#F4CCCC")
        .setRanges([statusRange])
        .build());
    
    // Apply all rules
    auditSheet.setConditionalFormatRules(rules);
    
    // Store last format time
    props.setProperty('lastAuditFormatTime', now.toString());
}

/**
 * Archive old audit log data to prevent sheet from growing too large
 */
function archiveOldData() {
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) {
            console.log("Could not obtain lock for archiving. Another process is running.");
            return;
        }
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
        
        if (!auditLogSheet || auditLogSheet.getLastRow() <= 5000) {
            return; // No need to archive
        }
        
        // Create archive sheet if it doesn't exist
        let archiveSheet = sheet.getSheetByName("Audit Archive");
        if (!archiveSheet) {
            archiveSheet = sheet.insertSheet("Audit Archive");
            // Copy headers
            auditLogSheet.getRange(1, 1, 1, auditLogSheet.getLastColumn())
                .copyTo(archiveSheet.getRange(1, 1));
        }
        
        // Get oldest 1000 entries (after header)
        const numRows = 1000;
        const dataToArchive = auditLogSheet.getRange(
            2, 1, numRows, auditLogSheet.getLastColumn()
        ).getValues();
        
        // Append to archive with retry
        retryOperation(() => {
            archiveSheet.getRange(
                archiveSheet.getLastRow() + 1, 
                1, 
                dataToArchive.length, 
                dataToArchive[0].length
            ).setValues(dataToArchive);
        });
        
        // Delete from audit log with retry
        retryOperation(() => {
            auditLogSheet.deleteRows(2, numRows);
        });
        
        // Clear processed responses cache
        CacheService.getScriptCache().remove('processedResponses');
        
        console.log(`✅ Archived ${numRows} audit log entries`);
    } catch (e) {
        console.error("❌ Error in archiveOldData:", e.message, e.stack);
        logError('Archive Data', `Error archiving data: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Add test data to queue for benchmarking
 */
function addTestData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const mnemonics = sheet.getSheetByName(SHEETS.SCORES).getRange('A2:A').getValues()
        .map(row => row[0])
        .filter(Boolean);
    
    // Create queue if it doesn't exist
    const queueSheet = sheet.getSheetByName("Processing Queue") || 
                        sheet.insertSheet("Processing Queue");
    
    // If queue sheet is new, add headers
    if (queueSheet.getLastRow() === 0) {
        queueSheet.appendRow(["Timestamp", "Mnemonic", "Processed", "Processing Timestamp"]);
    }
    
    // Add test entries - default to 10 entries for performance
    const testCount = Math.min(mnemonics.length, 10);
    const testRows = [];
    
    for (let i = 0; i < testCount; i++) {
        testRows.push([new Date(), mnemonics[i], "No", ""]);
    }
    
    if (testRows.length > 0) {
        queueSheet.getRange(
            queueSheet.getLastRow() + 1, 
            1, 
            testRows.length, 
            4
        ).setValues(testRows);
        
        console.log(`✅ Added ${testRows.length} test entries to queue`);
    }
}

/**
 * Retry operation with exponential backoff
 */
function retryOperation(operation, maxRetries = 3) {
    let retries = 0;
    while (retries < maxRetries) {
        try {
            return operation();
        } catch (e) {
            if (e.toString().includes("Rate Limit") || 
                e.toString().includes("Too many requests") ||
                e.toString().includes("exceeded maximum execution time") ||
                e.toString().includes("Service unavailable")) {
                
                if (retries < maxRetries - 1) {
                    const backoffTime = Math.pow(2, retries) * 1000 + Math.random() * 1000;
                    Utilities.sleep(backoffTime);
                    retries++;
                    console.warn(`Retry attempt ${retries} after ${backoffTime}ms delay...`);
                } else {
                    throw e;
                }
            } else {
                throw e;
            }
        }
    }
}

/**
 * Flush all queued items (admin function for immediate processing)
 */
function flushQueue() {
    const lock = LockService.getScriptLock();
    try {
        if (!lock.tryLock(10000)) {
            console.log("Could not obtain lock for flushing queue. Another process is running.");
            return;
        }
        
        console.log("🔄 Flushing all queued submissions...");
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        const queueSheet = sheet.getSheetByName("Processing Queue");
        
        if (!queueSheet) {
            console.log("ℹ️ No queue sheet found - nothing to process");
            return;
        }
        
        const queueData = queueSheet.getDataRange().getValues();
        if (queueData.length <= 1) {
            console.log("ℹ️ No items in queue to process");
            return;
        }
        
        // Find all pending indices
        const pendingIndices = [];
        for (let i = 1; i < queueData.length; i++) {
            if (queueData[i][2] === "No") {
                pendingIndices.push(i);
            }
        }
        
        if (pendingIndices.length === 0) {
            console.log("ℹ️ No pending items in queue");
            return;
        }
        
        console.log(`🔄 Flushing ${pendingIndices.length} queued submissions...`);
        
        // Process all pending items
        const BATCH_SIZE = 10;
        const startTime = new Date().getTime();
        const MAX_EXECUTION_TIME = 250000; // 250 seconds
        
        let processedCount = 0;
        
        for (let i = 0; i < pendingIndices.length; i += BATCH_SIZE) {
            // Check if approaching time limit
            if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
                console.warn(`⏱️ Approaching time limit. Processed ${processedCount} items.`);
                
                // Schedule continuation
                ScriptApp.newTrigger('flushQueue')
                    .timeBased()
                    .after(1000) // 1 second
                    .create();
                    
                return;
            }
            
            const batchIndices = pendingIndices.slice(i, i + BATCH_SIZE);
            processBatchFromQueue(batchIndices, queueSheet, queueData);
            processedCount += batchIndices.length;
        }
        
        console.log(`✅ Successfully flushed ${processedCount} queued submissions`);
        
        // Update processed responses
        updateProcessedResponses();
    } catch (e) {
        console.error("❌ Error in flushQueue:", e.message, e.stack);
        logError('Flush Queue', `Error flushing queue: ${e.message}\n${e.stack}`);
    } finally {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
    }
}

/**
 * Fix timestamp formatting in Form Responses (Raw) and other sheets
 */
function fixTimestampFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const rawResponsesSheet = sheet.getSheetByName("Form Responses (Raw)");
  const processedSheet = sheet.getSheetByName("Processed Responses");
  
  if (rawResponsesSheet && rawResponsesSheet.getLastRow() > 1) {
    // Get the format from an existing cell
    const existingFormat = rawResponsesSheet.getRange("A2").getNumberFormat();
    
    // Apply to all timestamp cells
    const timestampRange = rawResponsesSheet.getRange(2, 1, rawResponsesSheet.getLastRow()-1, 1);
    timestampRange.setNumberFormat(existingFormat);
  }
  
  if (processedSheet && processedSheet.getLastRow() > 1) {
    // Get the format from an existing cell if available, otherwise use a default format
    let existingFormat = "M/d/yyyy h:mm:ss";
    if (processedSheet.getLastRow() > 1) {
      existingFormat = processedSheet.getRange("A2").getNumberFormat() || existingFormat;
    }
    
    // Apply to all timestamp cells
    const timestampRange = processedSheet.getRange(2, 1, processedSheet.getLastRow()-1, 1);
    timestampRange.setNumberFormat(existingFormat);
    
    // Apply to processing timestamp column
    const processingTimestampRange = processedSheet.getRange(2, 4, processedSheet.getLastRow()-1, 1);
    processingTimestampRange.setNumberFormat(existingFormat);
  }
}


/**
 * Fix text alignment in all sheets
 */
function fixTextAlignment() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToFix = [
    "Form Responses (Raw)", 
    "Processed Responses",
    "Processing Queue"
  ];
  
  for (const sheetName of sheetsToFix) {
    const currentSheet = sheet.getSheetByName(sheetName);
    if (!currentSheet || currentSheet.getLastRow() <= 1) continue;
    
    // Get the last row and column
    const lastRow = currentSheet.getLastRow();
    const lastCol = currentSheet.getLastColumn();
    
    // Apply consistent alignment to all data cells
    currentSheet.getRange(2, 1, lastRow-1, lastCol)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
  }
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Quiz Admin')
        .addItem('Process Pending Responses', 'processQueue')
        .addItem('Flush All Pending Responses', 'flushQueue')
        .addItem('Grade Responses', 'gradeResponses')
        .addItem('Sync Responses', 'syncResponses')
        .addItem('Update Processed Responses', 'updateProcessedResponses')
        .addSeparator()
        .addItem('Delete Synced Data', 'deleteSyncData')
        .addItem('Reset Scores', 'resetAllScores')
        .addItem('Update Leaderboard', 'updateLeaderboard')
        .addSeparator()
        .addItem('Fix Formatting & Alignment', 'fixAllSheetFormatting')
        .addItem('Clear Audit Log', 'clearAuditLog')
        .addItem('Archive Old Audit Data', 'archiveOldData')
        .addItem('Add Test Data (10 Entries)', 'addTestData')
        .addItem('Clear Caches', 'clearCaches')
        .addSeparator()
        .addItem('Setup Automatic Processing', 'setupTriggers')
        .addItem('Update Audit Log Formatting', 'updateAuditLogFormatting')
        .addItem('Add Manual Points', 'showManualPointsDialog')
        .addItem('Process Manual Form Grades', 'handleFormGradeOverride')
        .addToUi();
}

/**
 * Fix all sheet formatting issues
 */
function fixAllSheetFormatting() {
  fixTimestampFormatting();
  fixTextAlignment();
  SpreadsheetApp.getActiveSpreadsheet().toast("Sheet formatting has been fixed", "Formatting Fixed", 5);
}

/**
 * Get answer letters for display
 */
function getAnswerLetters(answerText, qID, answerMapping) {
    if (!answerText || !answerMapping || !answerMapping[qID]) return answerText;

    if (answerText.includes(',')) {
        // Multiple select
        return answerText.split(',')
            .map(a => {
                const letterCode = answerMapping[qID][a.toLowerCase().trim()];
                return letterCode ? letterCode : a.trim();
            })
            .join(',');
    }

    // Single select
    const letterCode = answerMapping[qID][answerText.toLowerCase().trim()];
    return letterCode ? letterCode : answerText;
}

/**
 * Shorten long answer text for display
 */
function shortenAnswerText(answer, maxLength = 50) {
    if (!answer || answer.length <= maxLength) return answer;
    return answer.substring(0, maxLength) + "...";
}
