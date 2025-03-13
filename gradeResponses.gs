/**
 * Main grading function with optimizations, smart comma handling, and partial correctness
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
        
        // NEW: Get options map for smart comma handling
        const optionsMap = getOptionsMapWithCache(questionBankSheet);
        
        // Get valid mnemonics with caching
        const validMnemonics = getValidMnemonicsWithCache(scoresSheet);

        // Get only ungraded responses for efficiency
        const ungradedResponses = getUngradedResponses(responsesSheet, validMnemonics);
        
        if (ungradedResponses.length === 0) {
            console.info("✅ No ungraded responses to process");
            return;
        }
        
        console.info(`🔍 Processing ${ungradedResponses.length} ungraded responses`);

        // PERFORMANCE IMPROVEMENT: Build user data map once to avoid repeated scans
        const userDataMap = buildUserDataMap(scoresSheet);

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

            // Get user data from map instead of repeated sheet lookups
            const userData = userDataMap.get(mnemonic);
            if (!userData) {
                console.warn(`User ${mnemonic} not found in scores sheet`);
                continue;
            }

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

                // Get current score from user data map instead of scanning sheet
                const currentScore = userData.score;

                // Get actual role from user data map instead of scanning sheet
                const actualRole = (userData.role || "").trim().toLowerCase();
                const requiredRole = (questionData.targetRole || "").trim().toLowerCase();

                // Check both role mismatch & duplicate attempt at the same time
                const correctRole = actualRole === requiredRole || !requiredRole;
                
                // Check attempts from user data map
                const isDuplicate = hasAttemptInUserData(userData, qID);

                // UPDATED: Get options for this question for smart comma handling
                const questionOptions = optionsMap[qID] || [];

                // UPDATED: Variables to track correctness and partial credit
                const isMultipleSelect = questionData.type && questionData.type.toLowerCase() === "multiple select";
                let isCorrect = false;
                let isPartiallyCorrect = false;
                let earnedPoints = 0;

                // Grade the answer
                isCorrect = isAnswerCorrect(
                    userAnswer, 
                    questionData.correctAnswer, 
                    questionData.type,
                    questionOptions
                );

                // Only award points if eligible (correct role and not duplicate)
                if (correctRole && !isDuplicate) {
                    if (isMultipleSelect) {
                        // Calculate partial credit for multiple select
                        earnedPoints = calculatePartialCredit(
                            userAnswer,
                            questionData.correctAnswer,
                            questionData.type,
                            questionData.points,
                            questionOptions
                        );
                        
                        // Check if this is partially correct (some points but not full)
                        if (earnedPoints > 0 && earnedPoints < questionData.points) {
                            isPartiallyCorrect = true;
                        }
                    } else {
                        earnedPoints = isCorrect ? questionData.points : 0;
                    }

                    // Update scores and user data map
                    updateScores(scoresSheet, mnemonic, qID, earnedPoints, timestamp);
                    
                    // Update the in-memory score as well
                    userData.score += earnedPoints;
                    
                    // Add the attempt to userData
                    if (!userData.attempts) userData.attempts = {};
                    userData.attempts[qID] = { timestamp, points: earnedPoints };
                }

                // Determine correctness display status for audit log
                let correctnessStatus;
                if (isPartiallyCorrect) {
                    correctnessStatus = "Partially Correct";
                } else if (isCorrect) {
                    correctnessStatus = "Correct";
                } else {
                    correctnessStatus = "Incorrect";
                }

                // Update raw responses with correct/incorrect status
                // For the response sheet, we'll still use binary Correct/Incorrect
                responsesSheet.getRange(rowIndex + 1, 6).setValue(isCorrect || isPartiallyCorrect ? "Correct" : "Incorrect");

                // Get shortened answers for display
                let formattedUserAnswer = getAnswerLetters(userAnswer, qID, answerMapping);
                let formattedCorrectAnswer = getAnswerLetters(questionData.correctAnswer, qID, answerMapping);
                
                // Short display version for audit log
                const answerDisplay = `Answer: ${shortenAnswerText(formattedUserAnswer)} (Expected: ${shortenAnswerText(formattedCorrectAnswer)})`;

                // Log to audit with new column structure and updated correctness status
                auditLogEntries.push([
                    timestamp,                    // Timestamp
                    mnemonic,                    // Mnemonic
                    qID,                         // Question ID
                    answerDisplay,               // Shortened answer
                    correctnessStatus,           // UPDATED: Now shows "Partially Correct" when applicable
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

        // Update audit log formatting with new conditional formatting for "Partially Correct"
        updateAuditLogFormattingWithPartialCorrect();
        
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
 * Update audit log formatting, including special formatting for "Partially Correct"
 */
function updateAuditLogFormattingWithPartialCorrect() {
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
    
    // Correctness column (column E or 5)
    const correctnessColumn = 5;
    const correctnessRange = auditSheet.getRange(2, correctnessColumn, lastRow - 1, 1);
    
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
    
    // NEW: Format for correctness column
    // Correct - green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Correct")
        .setBackground("#D9EAD3")
        .setFontColor("#38761d")
        .setRanges([correctnessRange])
        .build());
    
    // Partially Correct - light green/yellow
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Partially Correct")
        .setBackground("#E7F1D7") // lighter green
        .setFontColor("#7F6000")  // dark yellow/gold
        .setRanges([correctnessRange])
        .build());
    
    // Incorrect - light red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Incorrect")
        .setBackground("#F4CCCC")
        .setFontColor("#990000")
        .setRanges([correctnessRange])
        .build());
    
    // Apply all rules
    auditSheet.setConditionalFormatRules(rules);
    
    // Store last format time
    props.setProperty('lastAuditFormatTime', now.toString());
}

/**
 * Helper function to build a map of user data for faster lookups
 */
function buildUserDataMap(scoresSheet) {
    const data = scoresSheet.getDataRange().getValues();
    const userMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) { // If mnemonic exists
            try {
                // Parse attempts JSON if it exists
                let attempts = {};
                if (data[i][5]) {
                    attempts = JSON.parse(data[i][5] || "{}");
                }
                
                userMap.set(data[i][0].toLowerCase(), {
                    score: Number(data[i][3]) || 0,  // Total score
                    role: data[i][2] || "",         // User role
                    rowIndex: i + 1,                // For updates later
                    attempts: attempts              // Parsed attempts
                });
            } catch (e) {
                console.error(`Error parsing attempts for ${data[i][0]}:`, e);
            }
        }
    }
    
    return userMap;
}

/**
 * Check if user has attempted question before using userData
 */
function hasAttemptInUserData(userData, questionID) {
    if (!userData.attempts) return false;
    return questionID in userData.attempts;
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
    
    // Clear the cache first to force refresh
    cache.remove(cacheKey);
    
    // Build the question map fresh
    const questionBankData = questionBankSheet.getDataRange().getValues();
    const questionMap = {};
    
    for (let i = 1; i < questionBankData.length; i++) {
        const row = questionBankData[i];
        const qID = row[1];
        if (qID) {
            // Log the correctAnswer for debugging
            console.log(`Question ${qID} correct answer: ${row[9]}`);
            
            questionMap[qID] = {
                question: row[2],
                correctAnswer: row[9],
                type: row[10],
                targetRole: row[11],
                points: parseInt(row[12]) || 0
            };
        }
    }
    
    return questionMap;
}

/**
 * Get answer mapping with caching
 */
/**
 * is function getAnswerMappingWithCache(questionMap) {
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
*/

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
/**function processQueue() {
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
*/

function processQueue() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
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
        
        // Only process a limited number at once to avoid timeout
        if (pendingCount >= 5) break;
      }
    }
    
    if (pendingCount === 0) {
      console.log("ℹ️ No pending items in queue");
      return;
    }
    
    console.log(`🔄 Processing ${pendingCount} queued submissions...`);
    
    // First sync all responses to make sure we have the latest data
    // But skip if we already have processed a lot recently
    if (pendingCount < 3) {
      syncResponses();
    }
    
    // Process limited batch to avoid timeout
    const batchIndices = pendingIndices.slice(0, pendingCount);
    processBatchFromQueue(batchIndices, queueSheet, queueData);
    
    // Only update processed responses if we're not about to time out
    updateProcessedResponses();
    
    // If more remain, schedule another run
    if (pendingCount < pendingIndices.length) {
      ScriptApp.newTrigger('processQueue')
        .timeBased()
        .after(60000) // 1 minute
        .create();
      console.log("⏱️ Scheduled additional queue processing in 1 minute");
    }
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
 * Grade responses for a specific mnemonic with smart comma handling
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
        const questionBankSheet = sheet.getSheetByName(SHEETS.QUESTION_BANK);
        const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
        const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
        
        if (!responsesSheet || !questionBankSheet || !scoresSheet || !auditLogSheet) {
            console.error("❌ Missing required sheets for grading");
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
        
        // Get processed responses
        const processedResponses = getProcessedResponsesWithCache(auditLogSheet);
        
        // Get question data
        const questionMap = getQuestionMapWithCache(questionBankSheet);
        const answerMapping = getAnswerMappingWithCache(questionMap);
        
        // NEW: Get options map for smart comma handling
        const optionsMap = getOptionsMapWithCache(questionBankSheet);
        
        let auditLogEntries = [];
        
        // Process each ungraded response for this mnemonic
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
                
                // UPDATED: Get options for this question for smart comma handling
                const questionOptions = optionsMap[qID] || [];
                
                // UPDATED: Variables to track correctness and partial credit
                const isMultipleSelect = questionData.type && questionData.type.toLowerCase() === "multiple select";
                let isCorrect = false;
                let isPartiallyCorrect = false;
                let earnedPoints = 0;
                
                // UPDATED: Grade answer using smart comma handling
                isCorrect = isAnswerCorrect(
                    userAnswer, 
                    questionData.correctAnswer, 
                    questionData.type,
                    questionOptions
                );
                
                // Only award points if eligible
                if (correctRole && !isDuplicate) {
                    if (isMultipleSelect) {
                        // UPDATED: Calculate partial credit using smart comma handling
                        earnedPoints = calculatePartialCredit(
                            userAnswer,
                            questionData.correctAnswer,
                            questionData.type,
                            questionData.points,
                            questionOptions
                        );
                        
                        // Check if this is partially correct (some points but not full)
                        if (earnedPoints > 0 && earnedPoints < questionData.points) {
                            isPartiallyCorrect = true;
                        }
                    } else {
                        earnedPoints = isCorrect ? questionData.points : 0;
                    }
                    
                    // Update scores
                    updateScores(scoresSheet, mnemonic, qID, earnedPoints, timestamp);
                }
                
                // Determine correctness display status for audit log
                let correctnessStatus;
                if (isPartiallyCorrect) {
                    correctnessStatus = "Partially Correct";
                } else if (isCorrect) {
                    correctnessStatus = "Correct";
                } else {
                    correctnessStatus = "Incorrect";
                }
                
                // Update raw responses with correct/incorrect status
                retryOperation(() => {
                    responsesSheet.getRange(rowIndex + 1, 6).setValue(isCorrect || isPartiallyCorrect ? "Correct" : "Incorrect");
                });
                
                // UPDATED: Get shortened answers for display with smart comma handling
                let formattedUserAnswer = getAnswerLetters(userAnswer, qID, answerMapping, questionOptions);
                let formattedCorrectAnswer = getAnswerLetters(questionData.correctAnswer, qID, answerMapping, questionOptions);
                
                // Short display version for audit log
                const answerDisplay = `Answer: ${shortenAnswerText(formattedUserAnswer)} (Expected: ${shortenAnswerText(formattedCorrectAnswer)})`;
                
                // Add to audit log entries
                auditLogEntries.push([
                    timestamp,                    // Timestamp
                    mnemonic,                    // Mnemonic
                    qID,                         // Question ID
                    answerDisplay,               // Shortened answer
                    correctnessStatus,           // UPDATED: Now shows "Partially Correct" when applicable
                    isDuplicate ? "Yes" : "No",  // Duplicate Attempt?
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
        updateAuditLogFormattingWithPartialCorrect();
        
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

/**
 * Creates a combined menu for all functions when the spreadsheet is opened
 */
/**
 * Creates a combined menu for all functions when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create Quiz Admin menu (main menu)
  const quizMenu = ui.createMenu('Quiz Admin')
    // Response Processing Section
    .addItem('Process Pending Responses', 'processQueue')
    .addItem('Flush All Pending Responses', 'flushQueue')
    .addItem('Grade Responses', 'gradeResponses')
    .addItem('Sync Responses', 'syncResponses')
    .addSeparator()
    
    // Grading Management
    .addItem('Regrade Specific Mnemonic', 'promptForRegrade')
    .addItem('Fix Q009 CDF Answers', 'fixQ009Answers')
    .addItem('Clear Question Cache', 'clearQuestionCache')
    .addSeparator()
    
    // Leaderboard & Scores Section
    .addItem('Update Leaderboard', 'updateLeaderboard')
    .addSeparator()
    
    // Tournament Management (NEW SECTION)
    .addItem('Advance to Round 2 (Top 8)', 'advanceToRoundTwo')
    .addItem('Update Round 2 Scores', 'updateRoundTwoScores')
    .addItem('Advance to Round 3 (Top 4)', 'advanceToRoundThree')
    .addItem('Update Round 3 Scores', 'updateRoundThreeScores')
    .addItem('Advance to Round 4 (Finals)', 'advanceToRoundFour')
    .addItem('Update Round 4 Scores', 'updateRoundFourScores')
    .addItem('Determine Champion', 'determineChampion')
    .addSeparator()
    
    // Manual Points Management
    .addItem('Add Manual Points', 'showManualPointsDialog')
    .addItem('Process Manual Form Grades', 'handleFormGradeOverride')
    .addItem('Setup Manual Grade Processing Log', 'setupManualGradeProcessingLog')
    .addItem('Recover Bonus Points', 'recoverBonusPoints')
    .addSeparator()
    
    // Question Management
    .addItem('Add New Question', 'showQuestionBankEditor')
    .addItem('Edit Existing Question', 'editExistingQuestion')
    .addItem('Update Daily Questions', 'updateDailyQuestions')
    .addItem('Reset Daily Questions Trigger', 'resetDailyQuestionsTrigger')
    .addSeparator()
    
    // Maintenance Section
    .addItem('Fix Timestamp Display', 'fixTimestampDisplay')
    .addItem('Check Timestamps', 'checkForTimestamps')
    .addItem('Clean Duplicate Responses', 'cleanupProcessedResponses')
    .addItem('Fix Formatting & Alignment', 'fixAllSheetFormatting')
    .addSeparator()
    
    // Competition Management
    .addItem('Determine Weekly Winners', 'testDetermineWinners')
    .addSeparator()
    
    // Administrative Section
    .addItem('Reset All Scores', 'resetAllScores')
    .addItem('Clear Audit Log', 'clearAuditLog')
    .addItem('Archive Old Audit Data', 'archiveOldData')
    .addSeparator()
    
    // Testing & Configuration
    .addItem('Add Test Data (10 Entries)', 'addTestData')
    .addItem('Clear Caches', 'clearCaches')
    .addItem('Setup Automatic Processing', 'setupTriggers');
  
  // Create Backup submenu
  const backupMenu = ui.createMenu('Backup')
    .addItem('Create Backup Now', 'createSpreadsheetBackup')
    .addItem('List All Backups', 'listBackups')
    .addSeparator()
    .addItem('Set Custom Backup Folder', 'setBackupFolder')
    .addItem('Reset to Default Backup Folder', 'resetBackupFolder')
    .addSeparator()
    .addItem('Schedule Daily Backups', 'createDailyBackupTrigger')
    .addItem('Schedule Weekly Backups', 'createWeeklyBackupTrigger')
    .addItem('Remove Backup Schedule', 'removeBackupTriggers');
  
  // Add the backup submenu to the Quiz Admin menu
  quizMenu.addSubMenu(backupMenu);
  
  // Add the main menu to the UI
  quizMenu.addToUi();
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
/**
 * function getAnswerLetters(answerText, qID, answerMapping) {
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
*/

/**
 * Shorten long answer text for display
 */
function shortenAnswerText(answer, maxLength = 50) {
    if (!answer || answer.length <= maxLength) return answer;
    return answer.substring(0, maxLength) + "...";
}

function restoreTimestamps() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const processedSheet = sheet.getSheetByName("Processed Responses");
  const rawResponsesSheet = sheet.getSheetByName("Form Responses (Raw)");
  
  if (!processedSheet || !rawResponsesSheet) {
    console.error("Required sheets not found");
    return;
  }
  
  // Get data from both sheets
  const processedData = processedSheet.getDataRange().getValues();
  const rawData = rawResponsesSheet.getDataRange().getValues();
  
  // Create a map of mnemonic to timestamps from processed responses
  const timestampMap = new Map();
  
  for (let i = 1; i < processedData.length; i++) {
    const timestamp = processedData[i][0]; // Timestamp is in column A
    const mnemonic = processedData[i][1]?.toLowerCase()?.trim() || ""; // Mnemonic is in column B
    
    if (timestamp && mnemonic) {
      timestampMap.set(mnemonic, timestamp);
    }
  }
  
  // Track updates needed
  let updatesNeeded = 0;
  const updates = [];
  
  // Check raw responses sheet for missing timestamps
  for (let i = 1; i < rawData.length; i++) {
    const existingTimestamp = rawData[i][0];
    const mnemonic = rawData[i][1]?.toLowerCase()?.trim() || "";
    
    // If timestamp is missing but we have mnemonic data
    if ((!existingTimestamp || existingTimestamp === "") && mnemonic) {
      const matchTimestamp = timestampMap.get(mnemonic);
      
      if (matchTimestamp) {
        // Queue the update
        updates.push([i+1, matchTimestamp]);
        updatesNeeded++;
      }
    }
  }
  
  console.log(`Found ${updatesNeeded} missing timestamps to restore`);
  
  if (updatesNeeded > 0) {
    for (const [row, timestamp] of updates) {
      rawResponsesSheet.getRange(row, 1).setValue(timestamp);
      // Add a small delay to prevent overloading
      Utilities.sleep(50);
    }
    
    // Fix formatting
    if (rawResponsesSheet.getLastRow() > 1) {
      const timestampRange = rawResponsesSheet.getRange(2, 1, rawResponsesSheet.getLastRow()-1, 1);
      timestampRange.setNumberFormat("M/d/yyyy h:mm:ss");
    }
    
    console.log(`Restored ${updatesNeeded} timestamps`);
  } else {
    console.log("No missing timestamps found");
  }
}

function cleanupProcessedResponses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const processedSheet = sheet.getSheetByName("Processed Responses");
  
  if (!processedSheet) {
    console.error("Processed Responses sheet not found");
    return;
  }
  
  const data = processedSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    console.log("No data to clean up");
    return;
  }
  
  // Create a set of unique entries
  const uniqueEntries = new Map();
  const duplicates = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const timestamp = data[i][0];
    const mnemonic = data[i][1];
    
    if (!timestamp || !mnemonic) continue;
    
    const key = `${mnemonic}_${timestamp}`;
    
    if (!uniqueEntries.has(key)) {
      uniqueEntries.set(key, i+1); // Store row number
    } else {
      duplicates.push(i+1); // This is a duplicate row
    }
  }
  
  console.log(`Found ${duplicates.length} duplicate entries`);
  
  // Delete duplicates in reverse order to avoid index shifting
  if (duplicates.length > 0) {
    for (let i = duplicates.length - 1; i >= 0; i--) {
      processedSheet.deleteRow(duplicates[i]);
    }
    console.log(`Removed ${duplicates.length} duplicate entries`);
  }
}

function diagnoseTimestampIssue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = sheet.getSheetByName("Form Responses (Raw)");
  const formSheet = sheet.getSheetByName(SHEETS.FORM_RESPONSES);
  
  if (!rawSheet || !formSheet) {
    console.log("Required sheets not found");
    return;
  }
  
  // Check protection
  const protections = rawSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (const p of protections) {
    const range = p.getRange();
    if (range.getColumn() === 1) {
      console.log("🔒 Column A is protected! Can't modify it.");
    }
  }
  
  // Check data validation
  const validation = rawSheet.getRange("A2").getDataValidation();
  if (validation) {
    console.log("⚠️ Column A has data validation rules: " + validation.getCriteriaType());
  }
  
  // Check for formulas
  const formulas = rawSheet.getRange("A2:A10").getFormulas();
  let hasFormulas = false;
  for (const row of formulas) {
    if (row[0]) {
      hasFormulas = true;
      console.log("📝 Column A contains formulas: " + row[0]);
      break;
    }
  }
  
  // Test writing a timestamp and check if it persists
  console.log("🧪 Testing timestamp writing...");
  const testRow = 2; // Use the second row for testing
  const testTimestamp = new Date();
  rawSheet.getRange(testRow, 1).setValue(testTimestamp);
  
  // Get the value immediately after setting
  const immediateValue = rawSheet.getRange(testRow, 1).getValue();
  console.log("📊 Immediate value after setting: " + immediateValue);
  
  // Wait a moment and check again
  Utilities.sleep(2000);
  const delayedValue = rawSheet.getRange(testRow, 1).getValue();
  console.log("📊 Value after 2 second delay: " + delayedValue);
  
  // Try writing with different format
  console.log("🧪 Testing with different format...");
  const formattedTimestamp = Utilities.formatDate(testTimestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  rawSheet.getRange(testRow, 1).setValue(formattedTimestamp);
  
  Utilities.sleep(2000);
  const finalValue = rawSheet.getRange(testRow, 1).getValue();
  console.log("📊 Final value after formatted write: " + finalValue);
  
  // Check display values vs actual values
  const displayValue = rawSheet.getRange(testRow, 1).getDisplayValue();
  console.log("📝 Display value: " + displayValue);
  const actualValue = rawSheet.getRange(testRow, 1).getValue();
  console.log("📝 Actual value: " + actualValue);
  
  // Check sheet triggers
  const triggers = ScriptApp.getProjectTriggers();
  console.log("⚙️ Checking for triggers that might interfere:");
  for (const trigger of triggers) {
    console.log(" - " + trigger.getHandlerFunction() + " (event: " + trigger.getEventType() + ")");
  }
  
  // Check if column is hidden
  const isHidden = rawSheet.isColumnHidden(1);
  console.log("👀 Column A hidden? " + isHidden);
  
  // Get sample timestamps from form responses
  const formData = formSheet.getRange("A2:A10").getValues();
  console.log("📅 Sample timestamps from Form Responses:");
  for (const row of formData) {
    if (row[0]) {
      console.log(" - " + row[0] + " (type: " + typeof row[0] + ")");
    }
  }
}

function fixTimestampDisplay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = sheet.getSheetByName("Form Responses (Raw)");
  
  if (!rawSheet) {
    console.log("Form Responses (Raw) sheet not found");
    return;
  }
  
  // Get the last row with data
  const lastRow = rawSheet.getLastRow();
  if (lastRow <= 1) {
    console.log("No data rows to fix");
    return;
  }
  
  // Try different date formats to see which one works
  console.log("Applying different date formats to make timestamps visible...");
  
  // First format - standard date/time
  rawSheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  
  // Second approach - directly modify the column format
  const sheet_id = sheet.getSheetId();
  const column_index = 1; // Column A
  
  try {
    const request = {
      'requests': [
        {
          'updateDimensionProperties': {
            'properties': {
              'numberFormat': {
                'type': 'DATE_TIME',
                'pattern': 'M/d/yyyy h:mm:ss'
              }
            },
            'fields': 'numberFormat',
            'range': {
              'sheetId': sheet_id,
              'dimension': 'COLUMNS',
              'startIndex': column_index - 1,
              'endIndex': column_index
            }
          }
        }
      ]
    };
    
    Sheets.Spreadsheets.batchUpdate(request, sheet.getId());
    console.log("Applied column format via API");
  } catch (e) {
    console.log("API format update failed: " + e.message);
    // Fall back to direct formatting
    rawSheet.getRange("A:A").setNumberFormat("M/d/yyyy h:mm:ss");
    console.log("Applied fallback formatting method");
  }
  
  // Try refreshing the values by writing them back
  const timestampValues = rawSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  rawSheet.getRange(2, 1, lastRow - 1, 1).setValues(timestampValues);
  console.log("Refreshed timestamp values");
}

/**
 * Prompts user for mnemonic and question ID to regrade
 */
function promptForRegrade() {
  const ui = SpreadsheetApp.getUi();
  
  // First prompt for mnemonic
  const mnemonicResponse = ui.prompt(
    'Regrade Student Answer',
    'Enter the student mnemonic to regrade:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (mnemonicResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const mnemonic = mnemonicResponse.getResponseText().trim();
  if (!mnemonic) {
    ui.alert('Error', 'Mnemonic cannot be empty', ui.ButtonSet.OK);
    return;
  }
  
  // Then prompt for question ID
  const questionResponse = ui.prompt(
    'Question to Regrade',
    'Enter the question ID to regrade (e.g., Q0009):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (questionResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const questionID = questionResponse.getResponseText().trim().toUpperCase();
  if (!questionID) {
    ui.alert('Error', 'Question ID cannot be empty', ui.ButtonSet.OK);
    return;
  }
  
  // Show progress dialog
  const progressMsg = ui.alert(
    'Processing',
    `Starting to regrade ${mnemonic} for question ${questionID}. Click OK to continue.`,
    ui.ButtonSet.OK
  );
  
  // Run the regrade operation
  try {
    const result = regradeSpecificAnswer(mnemonic, questionID);
    
    // Show result
    if (result.success) {
      ui.alert(
        'Regrade Complete',
        `Successfully regraded ${mnemonic} for question ${questionID}.\n\n${result.message}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Regrade Error',
        `Error: ${result.message}`,
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    ui.alert(
      'Error Occurred',
      `An error occurred while regrading: ${e.message}`,
      ui.ButtonSet.OK
    );
  }
}


/**
 * Regrades a specific answer for a specific mnemonic and question
 * Ignores duplicate check since this is a manual override
 * @param {string} mnemonic - The student mnemonic
 * @param {string} questionID - The question ID to regrade
 * @returns {Object} Success status and message
 */
function regradeSpecificAnswer(mnemonic, questionID) {
  // Normalize inputs
  mnemonic = mnemonic.toLowerCase().trim();
  questionID = questionID.toUpperCase().trim();
  
  console.log(`🔄 Starting regrade for ${mnemonic} question ${questionID}`);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const responsesSheet = sheet.getSheetByName("Form Responses (Raw)");
    const questionBankSheet = sheet.getSheetByName(SHEETS.QUESTION_BANK);
    const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
    const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
    
    if (!responsesSheet || !questionBankSheet || !scoresSheet || !auditLogSheet) {
      return { success: false, message: "Required sheets not found" };
    }
    
    // Clear cache to ensure fresh data
    clearQuestionCache();
    
    // Get question data
    const questionBankData = questionBankSheet.getDataRange().getValues();
    let questionData = null;
    
    for (let i = 1; i < questionBankData.length; i++) {
      if (questionBankData[i][1] === questionID) {
        questionData = {
          question: questionBankData[i][2],
          correctAnswer: questionBankData[i][9],
          type: questionBankData[i][10],
          targetRole: questionBankData[i][11],
          points: parseInt(questionBankData[i][12]) || 0
        };
        break;
      }
    }
    
    if (!questionData) {
      return { success: false, message: `Question ${questionID} not found in question bank` };
    }
    
    console.log(`Question ${questionID} found with correct answer: ${questionData.correctAnswer}`);
    
    // Find the response for this mnemonic and question
    const responseData = responsesSheet.getDataRange().getValues();
    let rowIndex = -1;
    let answerData = null;
    let timestamp = null;
    
    for (let i = 1; i < responseData.length; i++) {
      const rowMnemonic = responseData[i][1]?.toLowerCase().trim();
      if (rowMnemonic === mnemonic) {
        const parsedAnswer = parseAnswer(responseData[i][2]);
        if (parsedAnswer && parsedAnswer[questionID]) {
          rowIndex = i;
          answerData = parsedAnswer;
          timestamp = responseData[i][0];
          break;
        }
      }
    }
    
    if (rowIndex === -1 || !answerData) {
      return { success: false, message: `No response found for ${mnemonic} and question ${questionID}` };
    }
    
    const userAnswer = answerData[questionID];
    console.log(`Found response for ${mnemonic}: ${userAnswer}`);
    
    // Get user role
    const userData = getUserData(scoresSheet, mnemonic);
    if (!userData) {
      return { success: false, message: `User ${mnemonic} not found in scores sheet` };
    }
    
    const actualRole = (userData.role || "").trim().toLowerCase();
    const requiredRole = (questionData.targetRole || "").trim().toLowerCase();
    
    // Check role match
    const correctRole = actualRole === requiredRole || !requiredRole;
    if (!correctRole) {
      return { 
        success: true, 
        message: `User ${mnemonic} has role ${actualRole} but question requires ${requiredRole}. No points awarded.` 
      };
    }
    
    // Check if already attempted (for informational purposes only)
    const isDuplicate = hasAttemptedBefore(scoresSheet, mnemonic, questionID);
    console.log(`User has attempted before: ${isDuplicate} (will be ignored for manual regrade)`);
    
    // Grade the answer
    const isCorrect = isAnswerCorrect(userAnswer, questionData.correctAnswer, questionData.type);
    
    // Update raw responses with correct/incorrect status
    const currentStatus = responsesSheet.getRange(rowIndex + 1, 6).getValue();
    responsesSheet.getRange(rowIndex + 1, 6).setValue(isCorrect ? "Correct" : "Incorrect");
    
    let earnedPoints = 0;
    let message = "";
    
    // Award points if answer is correct, ignoring duplicate status
    if (correctRole && isCorrect) {
      // Calculate points
      if (questionData.type && questionData.type.toLowerCase() === "multiple select") {
        earnedPoints = calculatePartialCredit(
          userAnswer, questionData.correctAnswer, questionData.type, questionData.points
        );
      } else {
        earnedPoints = questionData.points;
      }
      
      // Calculate new score
      const currentScore = getCurrentScore(scoresSheet, mnemonic);
      
      // If this is a duplicate attempt, we need to handle it specially to update the existing attempt
      if (isDuplicate) {
        // Remove previous attempt first
        removeAttempt(scoresSheet, mnemonic, questionID);
      }
      
      // Add the new attempt with points
      updateScores(scoresSheet, mnemonic, questionID, earnedPoints, timestamp);
      
      // Get new score
      const newScore = getCurrentScore(scoresSheet, mnemonic);
      
      message = `Answer changed from ${currentStatus} to ${isCorrect ? "Correct" : "Incorrect"}. ` +
                `Points awarded: ${earnedPoints}. ` +
                `Previous score: ${currentScore}, new score: ${newScore}. ` + 
                (isDuplicate ? "(Previous attempt was overridden)" : "");
    } else if (!isCorrect) {
      message = `Answer is incorrect. No points awarded.`;
    }
    
    // Log this regrade action in the audit log
    auditLogSheet.appendRow([
      new Date(),           // Timestamp
      mnemonic,            // Mnemonic
      questionID,          // Question ID
      userAnswer,          // Answer
      isCorrect ? "Correct" : "Incorrect",  // Correct?
      "No",               // Duplicate? (marked No for manual regrading)
      correctRole ? "Yes" : "No",  // Correct Role?
      userData.score,      // Previous Points
      earnedPoints,        // Earned Points
      userData.score + earnedPoints, // Total Points
      "Manual Regrade"     // Status
    ]);
    
    // Update leaderboard
    updateLeaderboard();
    
    return { success: true, message: message };
  } catch (e) {
    console.error(`Error regrading ${mnemonic} for ${questionID}:`, e);
    return { success: false, message: e.message };
  }
}

/**
 * Helper function to get user data from scores sheet
 */
function getUserData(scoresSheet, mnemonic) {
  const data = scoresSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toLowerCase().trim() === mnemonic.toLowerCase().trim()) {
      return {
        mnemonic: data[i][0],
        name: data[i][1],
        role: data[i][2],
        score: Number(data[i][3]) || 0,
        row: i + 1
      };
    }
  }
  
  return null;
}

/**
 * Remove a previous attempt from a user's record
 */
function removeAttempt(scoresSheet, mnemonic, questionID) {
  const userData = getUserData(scoresSheet, mnemonic);
  if (!userData) return false;
  
  try {
    // Get existing attempts
    const attemptsCell = scoresSheet.getRange(userData.row, 6);
    const attemptsJson = attemptsCell.getValue();
    const attempts = JSON.parse(attemptsJson || "{}");
    
    // Check if the question is in attempts
    if (!(questionID in attempts)) return false;
    
    // Get previous points awarded for this question
    const previousPoints = attempts[questionID].points || 0;
    
    // Update user's score by subtracting previous points
    if (previousPoints > 0) {
      const scoreCell = scoresSheet.getRange(userData.row, 4);
      const currentScore = scoreCell.getValue();
      scoreCell.setValue(currentScore - previousPoints);
    }
    
    // Remove the question from attempts
    delete attempts[questionID];
    
    // Save updated attempts
    attemptsCell.setValue(JSON.stringify(attempts));
    
    return true;
  } catch (e) {
    console.error(`Error removing attempt for ${mnemonic} question ${questionID}:`, e);
    return false;
  }
}

/**
 * Helper function to get user data from scores sheet
 */
function getUserData(scoresSheet, mnemonic) {
  const data = scoresSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toLowerCase().trim() === mnemonic.toLowerCase().trim()) {
      return {
        mnemonic: data[i][0],
        name: data[i][1],
        role: data[i][2],
        score: Number(data[i][3]) || 0,
        row: i + 1
      };
    }
  }
  
  return null;
}

/**
 * Clears the cached question data to force a fresh load
 * This is useful when you've updated the Question Bank but changes aren't reflecting
 */
function clearQuestionCache() {
  const cache = CacheService.getScriptCache();
  
  // Remove all cached data related to questions
  cache.remove('questionMap');
  cache.remove('answerMapping');
  cache.remove('optionsMap');
  cache.remove('validMnemonics');
  cache.remove('processedResponses');
  
  // Let the user know it's done
  SpreadsheetApp.getUi().alert(
    'Cache Cleared',
    'Question cache has been cleared. The next grading operation will use fresh data from the Question Bank.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  console.log("🧹 Question cache cleared");
}

/**
 * Direct fix for Q009 answers that should be correct
 * This function directly fixes the issue without complex logic
 */
function fixQ009Answers() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask for mnemonic to fix
  const response = ui.prompt(
    'Fix Q009 Answer',
    'Enter the mnemonic to fix (or leave blank to fix all CDF answers):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const targetMnemonic = response.getResponseText().trim().toLowerCase();
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const responsesSheet = sheet.getSheetByName("Form Responses (Raw)");
  const auditLogSheet = sheet.getSheetByName(SHEETS.AUDIT_LOG);
  const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
  
  if (!responsesSheet || !auditLogSheet || !scoresSheet) {
    ui.alert("Error", "Required sheets not found", ui.ButtonSet.OK);
    return;
  }
  
  // Process the audit log
  const auditData = auditLogSheet.getDataRange().getValues();
  let fixedEntries = 0;
  
  for (let i = 1; i < auditData.length; i++) {
    const row = auditData[i];
    const mnemonic = row[1]?.toLowerCase();
    const questionID = row[2];
    const answerText = row[3];
    
    // Skip if not matching our target or not Q009
    if ((targetMnemonic && mnemonic !== targetMnemonic) || questionID !== "Q009") {
      continue;
    }
    
    // If answer contains C,D,F and expected contains C,D,F, mark as correct
    if (answerText.includes("Answer: C,D,F") && 
        answerText.includes("Expected: C")) {
      
      console.log(`Fixing row ${i+1} for ${mnemonic}`);
      
      // Mark as correct in audit log
      auditLogSheet.getRange(i+1, 5).setValue("Correct");
      
      // Award points (2 points for Q009)
      const currentScore = getCurrentScore(scoresSheet, mnemonic);
      updateScores(scoresSheet, mnemonic, "Q009", 2, new Date());
      
      // Mark as manually fixed
      auditLogSheet.getRange(i+1, 11).setValue("Manually Fixed");
      
      fixedEntries++;
    }
  }
  
  // Also fix in the raw responses sheet
  if (targetMnemonic) {
    const rawData = responsesSheet.getDataRange().getValues();
    
    for (let i = 1; i < rawData.length; i++) {
      const mnemonic = rawData[i][1]?.toLowerCase();
      
      if (mnemonic === targetMnemonic) {
        const answerData = parseAnswer(rawData[i][2]);
        
        if (answerData && answerData["Q009"] === "C,D,F") {
          responsesSheet.getRange(i+1, 6).setValue("Correct");
        }
      }
    }
  }
  
  // Update leaderboard
  updateLeaderboard();
  
  ui.alert("Fix Applied", `Fixed ${fixedEntries} entries for Q009.`, ui.ButtonSet.OK);
}