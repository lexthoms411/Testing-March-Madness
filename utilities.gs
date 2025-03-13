// Form and Sheet Configuration
const FORM_ID = "1btWSfeXXQJvh7wT8z9-Sk_d2pjeyLEf55dMWk63XsBo";

const SHEETS = {
    QUESTION_BANK: 'Question Bank',
    TEAMS: 'Teams',
    SCORES: 'Scores',
    TEAM_LEADERBOARD: 'Team Leaderboard',
    INDIVIDUAL_LEADERBOARD: 'Individual Leaderboard',
    AUDIT_LOG: 'Audit Log',
    FORM_RESPONSES: 'Form Responses',
    ERROR_LOG: 'Error Log'  // Added Error Log sheet
};

const SECTIONS = {
    RN: 'RN Section',
    PCA: 'PCA Section'
};

/**
 * Simple logging function
 */
function logToSheet(action, status, details) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const logSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

        if (logSheet) {
            logSheet.appendRow([
                new Date(),
                action,
                status,
                details
            ]);
        }
    } catch (e) {
        console.error('Logging failed:', e.message);
        logError('Sheet Logging', e.message);
    }
}

/**
 * Log error to Error Log sheet
 */
function logError(action, message, details = '') {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet();
        let errorSheet = sheet.getSheetByName(SHEETS.ERROR_LOG);
        
        if (!errorSheet) {
            errorSheet = sheet.insertSheet(SHEETS.ERROR_LOG);
            errorSheet.appendRow(['Timestamp', 'Action', 'Error Message', 'Details', 'Status']);
            errorSheet.setFrozenRows(1);
        }
        
        errorSheet.appendRow([
            new Date(),
            action,
            message,
            details,
            'Unresolved'
        ]);
    } catch (error) {
        console.error('Failed to log error:', error.message);
    }
}

/**
 * Mark error as resolved
 */
function resolveError(rowIndex) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = sheet.getSheetByName(SHEETS.ERROR_LOG);
    
    if (errorSheet) {
        errorSheet.getRange(rowIndex, 5).setValue('Resolved');
    }
}

/**
 * Parse answer JSON from form response
 */
function parseAnswer(answerString) {
    try {
        return JSON.parse(answerString);
    } catch (error) {
        console.error("❌ Error parsing answer JSON:", answerString, error);
        logError('Parse Answer', `Failed to parse answer: ${error.message}`, answerString);
        return {};
    }
}

/**
 * Normalize answer string for comparison - ENHANCED VERSION Smart Comma Separated
 */
function normalizeAnswer(answer, isMultipleSelect = false, options = []) {
  if (!answer) return '';

  let normalized = answer.toString()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

  if (isMultipleSelect && options.length > 0) {
    // Use smart processing
    return processMultipleSelectAnswer(normalized, options)
      .sort()  // Sort for consistent comparison
      .join(',');
  } else if (isMultipleSelect) {
    // Fall back to simple comma splitting if no options
    return normalized
      .split(',')
      .map(item => item.trim())
      .filter(Boolean)
      .sort()
      .join(',');
  }

  return normalized;
}


/**
 * Check if an answer is correct using smart comma handling
 * UPDATED: Correctly uses processMultipleSelectAnswer for option matching
 */
function isAnswerCorrect(userAnswer, correctAnswers, questionType, questionOptions = []) {
    if (!userAnswer || !correctAnswers) return false;

    const isMultipleSelect = questionType && questionType.toLowerCase() === "multiple select";

    // For debugging
    console.log(`Checking answer correctness for ${isMultipleSelect ? 'multiple select' : 'single select'} question`);
    console.log(`User answer: ${userAnswer}`);
    console.log(`Correct answer: ${correctAnswers}`);
    console.log(`Options available: ${questionOptions.length > 0 ? 'Yes' : 'No'}`);

    if (isMultipleSelect) {
        // Use smart processing for multiple select questions
        
        // First process the user's answer using smart option matching
        const userSelectedOptions = processMultipleSelectAnswer(userAnswer, [], questionOptions);
        console.log(`Smart processed user options: ${JSON.stringify(userSelectedOptions)}`);
        
        // Now process the correct answers the same way
        const correctSelectedOptions = processMultipleSelectAnswer(correctAnswers, [], questionOptions);
        console.log(`Smart processed correct options: ${JSON.stringify(correctSelectedOptions)}`);
        
        // Compare the sets - make sure all expected options are selected and no extras
        if (userSelectedOptions.length !== correctSelectedOptions.length) {
            console.log(`Length mismatch - user has ${userSelectedOptions.length} options, expected ${correctSelectedOptions.length}`);
            return false;
        }
        
        // Check if all user selections are in the correct answers
        const allUserSelectionsCorrect = userSelectedOptions.every(option => {
            const found = correctSelectedOptions.some(correctOption => 
                correctOption.toLowerCase().trim() === option.toLowerCase().trim());
            if (!found) console.log(`User selection "${option}" not found in correct options`);
            return found;
        });
        
        // Check if all correct answers are in the user selections
        const allCorrectSelectionsPresent = correctSelectedOptions.every(option => {
            const found = userSelectedOptions.some(userOption => 
                userOption.toLowerCase().trim() === option.toLowerCase().trim());
            if (!found) console.log(`Expected option "${option}" not selected by user`);
            return found;
        });
        
        // Both conditions must be true for the answer to be fully correct
        return allUserSelectionsCorrect && allCorrectSelectionsPresent;
    } else {
        // For single select questions, just compare the normalized text
        const normalizedUserAnswer = normalizeAnswer(userAnswer);
        const normalizedCorrectAnswer = normalizeAnswer(correctAnswers);
        return normalizedUserAnswer === normalizedCorrectAnswer;
    }
}


/**
 * Extract question ID from question text
 */
function extractQuestionID(questionText) {
    if (!questionText) return null;
    const match = questionText.match(/\[Q\d+\]/);
    return match ? match[0].replace(/\[|\]/g, '') : null;
}

/**
 * Check if user has attempted question before
 */
function hasAttemptedBefore(scoresSheet, mnemonic, questionID) {
    const attemptsRange = scoresSheet.getRange('F2:F').getValues();
    const mnemonics = scoresSheet.getRange('A2:A').getValues();

    for (let i = 0; i < mnemonics.length; i++) {
        if (mnemonics[i][0]?.toLowerCase() === mnemonic.toLowerCase()) {
            try {
                let attempts = JSON.parse(attemptsRange[i][0] || "{}");
                return questionID in attempts;
            } catch (e) {
                console.error(`Error checking attempts for ${mnemonic}:`, e);
                logError('Check Attempts', e.message, `Mnemonic: ${mnemonic}, QuestionID: ${questionID}`);
                return false;
            }
        }
    }
    return false;
}

/**
 * Get set of already processed responses
 */
function getProcessedResponses(auditLogSheet) {
    const processedResponses = new Set();
    const auditData = auditLogSheet.getDataRange().getValues();

    for (let i = 1; i < auditData.length; i++) {
        const key = `${auditData[i][0]}_${auditData[i][1]}_${auditData[i][2]}`.toLowerCase();
        processedResponses.add(key);
    }

    return processedResponses;
}

/**
 * Get the actual user role from the Scores sheet
 */
function getUserRole(scoresSheet, mnemonic) {
    const scoresData = scoresSheet.getDataRange().getValues();

    for (let i = 1; i < scoresData.length; i++) {
        if (scoresData[i][0]?.toLowerCase() === mnemonic.toLowerCase()) {
            return scoresData[i][2] || "";  // Column C contains the actual role (RN/PCA)
        }
    }

    console.warn(`⚠️ Role not found for ${mnemonic}, defaulting to empty role`);
    logError('Get User Role', `Role not found for ${mnemonic}`);
    return "";
}

/**
 * Get current score for a user
 */
function getCurrentScore(scoresSheet, mnemonic) {
    const data = scoresSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (data[i][0]?.toLowerCase() === mnemonic.toLowerCase()) {
            return Number(data[i][3]) || 0; // Column D contains the score
        }
    }
    return 0;
}

/**
 * Process multiple select answers with smart comma handling
 * UPDATED: More robust implementation that checks full option text
 */
/**
 * Process multiple select answers with improved comma handling for complex options
 */
function processMultipleSelectAnswer(answerText, options, questionOptions = []) {
    // Return empty array for empty input
    if (!answerText) return [];
    
    // Debug
    console.log(`Processing answer: ${answerText}`);
    
    // If we don't have options, fall back to simple comma splitting
    if ((!options || options.length === 0) && (!questionOptions || questionOptions.length === 0)) {
        console.log("No options available, using basic comma splitting");
        return answerText.split(',').map(item => item.trim()).filter(Boolean);
    }
    
    // Use questionOptions if provided (from optionsMap), else use options parameter
    const useOptions = (questionOptions && questionOptions.length > 0) ? questionOptions : options;
    
    if (useOptions.length === 0) {
        console.log("Warning: No usable options found, falling back to basic splitting");
        return answerText.split(',').map(item => item.trim()).filter(Boolean);
    }
    
    console.log(`Using ${useOptions.length} options for processing`);
    
    // Smart parsing using the known options
    const selectedOptions = [];
    
    // Normalize answerText - critical for consistent matching
    let normalizedAnswer = answerText.toLowerCase().trim();
    
    // Sort options by length (longest first) to avoid partial matches
    const sortedOptions = [...useOptions]
        .filter(Boolean) // Remove empty/null values
        .sort((a, b) => String(b).length - String(a).length);
    
    // Create a copy of the answer text for tracking what's been matched
    let remainingText = normalizedAnswer;
    
    // For each possible option, check if it appears in the answer
    for (const option of sortedOptions) {
        if (!option) continue;
        
        // Normalize option text
        const optionText = String(option).toLowerCase().trim();
        
        // Skip empty options
        if (!optionText) continue;
        
        // Different ways the option might appear in text
        const variations = [
            optionText,                     // Exact match
            optionText + ',',               // With trailing comma
            ', ' + optionText,              // With leading comma and space
            ',' + optionText,               // With leading comma
            ' ' + optionText + ' ',         // With spaces
            ' ' + optionText + ','          // With space and trailing comma
        ];
        
        // Check if any variation appears in the remaining text
        let foundMatch = false;
        for (const variant of variations) {
            if (remainingText.includes(variant)) {
                selectedOptions.push(option);
                console.log(`Found option: "${option}"`);
                
                // Remove this option from the remaining text
                remainingText = remainingText.replace(variant, ',');
                foundMatch = true;
                break;
            }
        }
        
        // If no exact match found, try a more flexible approach (for complex options)
        if (!foundMatch) {
            // For complex options, first try to check if all words are present
            const optionWords = optionText.split(/\s+/);
            const answerWords = remainingText.split(/\s+|,/);
            
            const allWordsPresent = optionWords.every(word => 
                answerWords.some(answerWord => answerWord === word));
                
            if (allWordsPresent && optionWords.length > 2) {
                // If we have at least 3 words and all are present, consider it a match
                selectedOptions.push(option);
                console.log(`Found option by word matching: "${option}"`);
                
                // Remove words from remaining text
                for (const word of optionWords) {
                    remainingText = remainingText.replace(word, '');
                }
            }
        }
    }
    
    // Clean up remaining text to find any additional entries
    remainingText = remainingText.replace(/\s+/g, ' ').trim();
    
    // If there's leftover text with commas, try to handle it
    if (remainingText && remainingText.includes(',')) {
        const remainingItems = remainingText
            .split(',')
            .map(item => item.trim())
            .filter(Boolean);
        
        if (remainingItems.length > 0) {
            console.log(`Found ${remainingItems.length} additional items in remaining text`);
            for (const item of remainingItems) {
                if (item.length > 3) { // Only add non-trivial items (more than 3 chars)
                    selectedOptions.push(item);
                }
            }
        }
    } else if (remainingText && remainingText.length > 3) {
        // Single leftover item (no commas)
        selectedOptions.push(remainingText);
    }
    
    // Add a final deduplication step - get unique options
    const uniqueOptions = [...new Set(selectedOptions.map(opt => String(opt).trim()))];
    
    console.log(`Final processed options: ${JSON.stringify(uniqueOptions)}`);
    return uniqueOptions;
}

/**
 * Helper function to escape regex special characters
 */
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Calculate partial credit for multiple select questions - FIXED
 */
/**
 * Calculate partial credit for multiple select questions - IMPROVED
 */
function calculatePartialCredit(userAnswer, correctAnswer, questionType, totalPoints, questionOptions = []) {
    if (questionType.toLowerCase() !== "multiple select") return 0;

    // Debug logging
    console.log(`Checking answer correctness for multiple select question`);
    console.log(`User answer: ${userAnswer}`);
    console.log(`Correct answer: ${correctAnswer}`);
    console.log(`Options available: ${questionOptions && questionOptions.length > 0 ? 'Yes' : 'No'}`);
    
    // Process options using smart comma handling
    const userItems = processMultipleSelectAnswer(userAnswer, [], questionOptions);
    const correctItems = processMultipleSelectAnswer(correctAnswer, [], questionOptions);
    
    // Debug the processed options
    console.log(`Smart processed user options: ${JSON.stringify(userItems)}`);
    console.log(`Smart processed correct options: ${JSON.stringify(correctItems)}`);

    // Check if we have valid data to proceed
    if (!correctItems || correctItems.length === 0) {
        console.log(`Warning: No correct items identified, cannot calculate partial credit`);
        return 0;
    }
    
    // Normalize and map to lowercase for comparison
    const userItemsLower = userItems.map(item => String(item).toLowerCase().trim());
    const correctItemsLower = correctItems.map(item => String(item).toLowerCase().trim());
    
    // Calculate correct selections (intersection of user selections and correct selections)
    const correctSelections = userItemsLower.filter(item => 
        correctItemsLower.some(correct => 
            item === correct || correct.includes(item) || item.includes(correct)));
    
    // Calculate incorrect selections (user selections not in correct selections)
    const incorrectSelections = userItemsLower.filter(item => 
        !correctItemsLower.some(correct => 
            item === correct || correct.includes(item) || item.includes(correct)));
    
    const correctCount = correctSelections.length;
    const incorrectCount = incorrectSelections.length;
    const totalCorrectItems = correctItemsLower.length;
    
    console.log(`Correct selections: ${correctCount}/${totalCorrectItems}`);
    console.log(`Incorrect selections: ${incorrectCount}`);
    
    // Guard against division by zero
    if (totalCorrectItems === 0) {
        console.log(`Error: Total correct items is zero, cannot calculate points per item`);
        return 0;
    }

    // Calculate points - award points for correct items, deduct for incorrect
    const pointsPerItem = totalPoints / totalCorrectItems;
    const earnedPoints = Math.max(0, Math.round((correctCount * pointsPerItem) - (incorrectCount * pointsPerItem)));
    
    console.log(`Partial credit: ${earnedPoints}/${totalPoints} (${correctCount} correct, ${incorrectCount} incorrect)`);
    
    return earnedPoints;
}

/**
 * Now we must update how getOptionsMapWithCache works to provide the actual options
 * This is different from the answerMapping which just maps to letters
 */
function getOptionsMapWithCache(questionBankSheet) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'optionsMap';
    
    // Try to get from cache first
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
        try {
            return JSON.parse(cachedData);
        } catch (e) {
            console.warn("⚠️ Cache parse error, rebuilding options map");
        }
    }
    
    // If not in cache or parse error, rebuild
    const optionsMap = getAnswerOptionsMap(questionBankSheet);
    
    // Cache for 6 hours
    try {
        cache.put(cacheKey, JSON.stringify(optionsMap), 21600);
    } catch (e) {
        console.warn("⚠️ Cache size limit exceeded, options not cached");
    }
    
    return optionsMap;
}


/**
 * Get all answer options for each question (full text)
 */
function getAnswerOptionsMap(questionBankSheet) {
    console.log("Building options map from question bank");
    
    const data = questionBankSheet.getDataRange().getValues();
    const optionsMap = {};
    
    for (let i = 1; i < data.length; i++) {
        const qID = data[i][1]; // Column B has question ID
        if (qID) {
            // Columns D-I (3-8) contain the answer options
            const options = [data[i][3], data[i][4], data[i][5], 
                            data[i][6], data[i][7], data[i][8]]
                            .filter(Boolean); // Remove empty entries
            
            // Store the full option text - essential for smart comma handling
            optionsMap[qID] = options;
            
            console.log(`Added ${options.length} options for ${qID}`);
        }
    }
    
    return optionsMap;
}

/**
 * Updated getAnswerLetters function to be more robust
 */
function getAnswerLetters(answerText, qID, answerMapping, questionOptions = []) {
    if (!answerText) return answerText;
    
    // Log for debugging
    console.log(`Processing answer for ${qID}: ${answerText}`);
    
    // Check if we have answer mapping for this question
    if (!answerMapping || !answerMapping[qID]) {
        console.log(`No letter mapping for ${qID}`);
        return answerText;
    }
    
    console.log(`Available mappings for ${qID}: ${JSON.stringify(answerMapping[qID] || {})}`);
    
    // For multiple select questions with commas, we need special handling
    if (answerText.includes(',')) {
        // Try to match complete option texts from the question options
        const fullOptions = questionOptions.length > 0 ? questionOptions : 
                            Object.keys(answerMapping[qID]);
        
        let result = [];
        // First, try to find exact matches for full options
        for (const option of fullOptions) {
            const optionLower = option.toLowerCase().trim();
            
            // Create different variations of the option text to check (with/without commas)
            const variations = [
                optionLower,
                optionLower + ',',
                ', ' + optionLower,
                ', ' + optionLower + ','
            ];
            
            // Check if any variation appears in the answer
            for (const variant of variations) {
                if (answerText.toLowerCase().includes(variant)) {
                    // If we find a match, look up the letter code
                    const letter = answerMapping[qID][optionLower];
                    if (letter) {
                        result.push(letter);
                        break;
                    }
                }
            }
        }
        
        // If we found letter mappings, use those
        if (result.length > 0) {
            return result.join(',');
        }
        
        // Fall back to basic comma splitting if we couldn't match full options
        return answerText.split(',')
            .map(a => {
                const key = a.toLowerCase().trim();
                const letter = answerMapping[qID][key];
                return letter || a.trim();
            })
            .join(',');
    } else {
        // Single select is simpler
        const key = answerText.toLowerCase().trim();
        const letter = answerMapping[qID][key];
        
        if (letter) {
            return letter;
        } else {
            console.log(`No mapping found for '${answerText}' in ${qID}`);
            return answerText;
        }
    }
}

/**
 * Get answer mapping with caching - IMPROVED
 */
function getAnswerMappingWithCache(questionMap) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'answerMapping';
    
    // Clear the cache to force refresh (remove this line for production)
    cache.remove(cacheKey);
    
    // Build the answer mapping fresh
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const questionBankSheet = sheet.getSheetByName(SHEETS.QUESTION_BANK);
    const questionBankData = questionBankSheet.getDataRange().getValues();
    const answerMapping = {};
    
    for (let i = 1; i < questionBankData.length; i++) {
        const row = questionBankData[i];
        const qID = row[1]; // Column B is question ID
        const type = row[10]; // Column K is question type
        
        if (qID && type && (type.toLowerCase().includes("multiple"))) {
            // Get options from columns D-I (index 3-8)
            const options = [row[3], row[4], row[5], row[6], row[7], row[8]].filter(Boolean);
            const letterMap = {};
            
            // Create mapping from option text to letter
            options.forEach((text, index) => {
                if (text) {
                    const letter = String.fromCharCode(65 + index); // A, B, C, etc.
                    letterMap[text.toLowerCase().trim()] = letter;
                }
            });
            
            // Only store if we have mappings
            if (Object.keys(letterMap).length > 0) {
                answerMapping[qID] = letterMap;
                
                // Debug output
                console.log(`Created letter mapping for ${qID}: ${JSON.stringify(letterMap)}`);
            }
        }
    }
    
    // Cache for 6 hours
    try {
        cache.put(cacheKey, JSON.stringify(answerMapping), 21600);
    } catch (e) {
        console.warn("Cache put error - likely too large:", e);
    }
    
    return answerMapping;
}
