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
 * Normalize answer string for comparison
 */
function normalizeAnswer(answer, isMultipleSelect = false) {
    if (!answer) return '';

    let normalized = answer.toString()
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();

    if (isMultipleSelect) {
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
 * Check if answer is correct
 */
function isAnswerCorrect(userAnswer, correctAnswers, questionType) {
    if (!userAnswer || !correctAnswers) return false;

    let isMultipleSelect = questionType.toLowerCase() === "multiple select";

    let normalizedUserAnswer = normalizeAnswer(userAnswer, isMultipleSelect);
    let correctAnswerList = correctAnswers.toString()
        .split(',')
        .map(ans => normalizeAnswer(ans, isMultipleSelect));

    if (isMultipleSelect) {
        let userAnswerSet = new Set(normalizedUserAnswer.split(','));
        let correctAnswerSet = new Set(correctAnswerList);
        return Array.from(userAnswerSet).sort().join(',') ===
               Array.from(correctAnswerSet).sort().join(',');
    }

    return correctAnswerList.some(correct =>
        correct === normalizedUserAnswer
    );
}

/**
 * Calculate partial credit for multiple select questions
 */
function calculatePartialCredit(userAnswer, correctAnswer, questionType, totalPoints) {
    if (questionType.toLowerCase() !== "multiple select") return 0;

    const userItems = new Set(userAnswer.toLowerCase().split(',').map(i => i.trim()));
    const correctItems = new Set(correctAnswer.toLowerCase().split(',').map(i => i.trim()));

    const correctCount = [...userItems].filter(item => correctItems.has(item)).length;
    const totalCorrectItems = correctItems.size;
    const incorrectCount = userItems.size - correctCount;

    const pointsPerItem = totalPoints / totalCorrectItems;
    return Math.max(0, Math.round((correctCount * pointsPerItem) - (incorrectCount * pointsPerItem)));
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
