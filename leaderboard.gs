/**
 * Update both leaderboards
 */
function updateLeaderboard() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const scoresSheet = sheet.getSheetByName(SHEETS.SCORES);
    const teamsSheet = sheet.getSheetByName(SHEETS.TEAMS);
    const teamLeaderboardSheet = sheet.getSheetByName(SHEETS.TEAM_LEADERBOARD);
    const individualLeaderboardSheet = sheet.getSheetByName(SHEETS.INDIVIDUAL_LEADERBOARD);
    const teamScoresSheet = sheet.getSheetByName('Team Scores');

    if (!scoresSheet || !teamsSheet || !teamLeaderboardSheet || !individualLeaderboardSheet || !teamScoresSheet) {
        console.error("❌ Required sheets not found for leaderboard update");
        return;
    }

    try {
        // Get all scores
        const scores = getScoresData(scoresSheet);
        
        // Update individual leaderboard
        updateIndividualLeaderboard(scores, individualLeaderboardSheet);

        // Get and sort team scores
        const teamScores = getTeamScoresData(teamScoresSheet);

        // Update team leaderboard
        updateTeamLeaderboard(teamScores, teamLeaderboardSheet);

    } catch (error) {
        console.error("❌ Error updating leaderboard:", error.message);
        logToSheet("Update Leaderboard", "ERROR", error.message);
    }
}

/**
 * Get scores data from scores sheet
 */
function getScoresData(scoresSheet) {
    const data = scoresSheet.getDataRange().getValues();
    const scores = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) { // If mnemonic exists
            scores.push({
                mnemonic: data[i][0].toLowerCase(),
                score: Number(data[i][3]) || 0  // Total score in column D
            });
        }
    }
    
    return scores;
}

/**
 * Get team data from teams sheet
 */
function getTeamData(teamsSheet) {
    const data = teamsSheet.getDataRange().getValues();
    const teams = new Map();
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][1]) { // If team name and mnemonic exist
            const teamName = data[i][0];
            const mnemonic = data[i][1].toLowerCase();
            
            if (!teams.has(teamName)) {
                teams.set(teamName, []);
            }
            teams.get(teamName).push(mnemonic);
        }
    }
    
    return teams;
}

/**
 * Get team scores from Team Scores sheet
 */
function getTeamScoresData(teamScoresSheet) {
    const data = teamScoresSheet.getDataRange().getValues();
    const teamScores = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) { // If team name exists
            teamScores.push({
                team: data[i][0],
                score: Number(data[i][1]) || 0,  // Total score in column B
                avgScore: 0  // Not used but kept for compatibility
            });
        }
    }
    
    // Sort by score descending
    return teamScores.sort((a, b) => b.score - a.score);
}


/**
 * Update individual leaderboard in its own sheet
 */
function updateIndividualLeaderboard(scores, leaderboardSheet) {
    // Clear existing data and formatting
    leaderboardSheet.clearConditionalFormatRules();
    
    // Get the last row, default to 1 if sheet is empty
    const lastRow = Math.max(leaderboardSheet.getLastRow(), 1);
    leaderboardSheet.getRange(1, 1, lastRow, 4).clear();  // Changed to 4 columns
    
    // Set headers
    leaderboardSheet.getRange('A1:D1')  // Changed to include Name column
        .setValues([['Rank', 'Mnemonic', 'Name', 'Total Score']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');

    // Get names from Scores sheet
    const scoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SCORES);
    const scoresData = scoresSheet.getDataRange().getValues();
    const nameMap = new Map();
    
    // Skip header row in scores sheet
    for (let i = 1; i < scoresData.length; i++) {
        if (scoresData[i][0]) {  // If mnemonic exists
            nameMap.set(scoresData[i][0].toLowerCase(), scoresData[i][1]);  // Map mnemonic to name
        }
    }

    // Sort scores by descending order
    const sortedScores = scores.sort((a, b) => b.score - a.score);
    
    // Create leaderboard data with ranks and names
    const leaderboardData = sortedScores.map((player, index) => [
        index + 1,
        player.mnemonic,
        nameMap.get(player.mnemonic) || '',  // Get name from map
        player.score
    ]);
    
    if (leaderboardData.length > 0) {
        leaderboardSheet.getRange(2, 1, leaderboardData.length, 4)
            .setValues(leaderboardData);
            
        // Format score column as numbers
        leaderboardSheet.getRange(2, 4, leaderboardData.length, 1)  // Changed to column D
            .setNumberFormat('#,##0')
            .setHorizontalAlignment('right');
    }

    // Add conditional formatting for top 3
    const rules = [];
    
    // Gold for 1st place
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([leaderboardSheet.getRange('A2:D2')])  // Changed to include Name column
        .whenFormulaSatisfied('=ROW()=2')
        .setBackground('#FFD700')
        .build());
    
    // Silver for 2nd place
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([leaderboardSheet.getRange('A3:D3')])  // Changed to include Name column
        .whenFormulaSatisfied('=ROW()=3')
        .setBackground('#C0C0C0')
        .build());
    
    // Bronze for 3rd place
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([leaderboardSheet.getRange('A4:D4')])  // Changed to include Name column
        .whenFormulaSatisfied('=ROW()=4')
        .setBackground('#CD7F32')
        .build());

    // Apply the rules
    leaderboardSheet.setConditionalFormatRules(rules);

    // Set column widths
    leaderboardSheet.setColumnWidth(1, 60);   // Rank column
    leaderboardSheet.setColumnWidth(2, 150);  // Mnemonic column
    leaderboardSheet.setColumnWidth(3, 200);  // Name column
    leaderboardSheet.setColumnWidth(4, 100);  // Score column
}

/**
 * Update team leaderboard
 */
function updateTeamLeaderboard(teamScores, leaderboardSheet) {
    // Clear existing data and formatting
    leaderboardSheet.clearConditionalFormatRules();
    
    // Get the last row, default to 1 if sheet is empty
    const lastRow = Math.max(leaderboardSheet.getLastRow(), 1);
    leaderboardSheet.getRange(1, 1, lastRow, 3).clear();
    
    // Set headers
    leaderboardSheet.getRange('A1:C1')
        .setValues([['Rank', 'Team', 'Total Score']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    
    // Create leaderboard data with ranks
    const leaderboardData = teamScores.map((team, index) => [
        index + 1,
        team.team,
        team.score
    ]);
    
    if (leaderboardData.length > 0) {
        leaderboardSheet.getRange(2, 1, leaderboardData.length, 3)
            .setValues(leaderboardData);
            
        // Format score column as numbers
        leaderboardSheet.getRange(2, 3, leaderboardData.length, 1)
            .setNumberFormat('#,##0')
            .setHorizontalAlignment('right');
    }

    // Add conditional formatting for top 3
    const rules = [];
    
    // Gold for 1st place
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([leaderboardSheet.getRange('A2:C2')])  // Changed: setRange instead of setRanges
        .whenFormulaSatisfied('=ROW()=2')
        .setBackground('#FFD700')
        .build());
    
    // Silver for 2nd place
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([leaderboardSheet.getRange('A3:C3')])  // Changed: setRange instead of setRanges
        .whenFormulaSatisfied('=ROW()=3')
        .setBackground('#C0C0C0')
        .build());
    
    // Bronze for 3rd place
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([leaderboardSheet.getRange('A4:C4')])  // Changed: setRange instead of setRanges
        .whenFormulaSatisfied('=ROW()=4')
        .setBackground('#CD7F32')
        .build());
        
    // Apply the rules
    leaderboardSheet.setConditionalFormatRules(rules);

    // Set column widths
    leaderboardSheet.setColumnWidth(1, 60);   // Rank column
    leaderboardSheet.setColumnWidth(2, 200);  // Team name column
    leaderboardSheet.setColumnWidth(3, 100);  // Score column
}

/**
 * Get current team rankings
 */
function getTeamRankings() {
    const leaderboardSheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEETS.LEADERBOARD);
    
    if (!leaderboardSheet) return [];
    
    const data = leaderboardSheet.getDataRange().getValues();
    const rankings = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            rankings.push({
                team: data[i][0],
                score: data[i][1]
            });
        }
    }
    
    return rankings;
}
