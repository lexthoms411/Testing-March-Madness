/**
 * Tournament advancement functions for March Madness challenge
 * Handles advancement for all rounds with date-based scoring
 */

// Round start dates
const ROUND_DATES = {
  ROUND_2: new Date("2025-03-09T00:00:00"),
  ROUND_3: new Date("2025-03-16T00:00:00"),
  ROUND_4: new Date("2025-03-23T00:00:00")
};

/**
 * Advances winners from Round 1 to Round 2
 */
function advanceToRoundTwo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamScoreSheet = ss.getSheetByName("Team Scores");
  
  // Determine winners from Round 1
  const winners = [];
  
  // Process each pair of teams (assuming 16 teams in Round 1)
  for (let i = 2; i <= 16; i += 2) {
    const team1 = teamScoreSheet.getRange(i, 1).getValue();
    const team2 = teamScoreSheet.getRange(i + 1, 1).getValue();
    const score1 = teamScoreSheet.getRange(i, 2).getValue() || 0;
    const score2 = teamScoreSheet.getRange(i + 1, 2).getValue() || 0;
    
    // Determine winner based on higher score
    const winner = (score1 >= score2) ? team1 : team2;
    const winnerScore = Math.max(score1, score2);
    
    // Calculate position for winner in Round 2
    // Each pair of teams in Round 1 produces one winner for Round 2
    const round2Row = Math.floor((i - 2) / 2) + 2;
    
    // Add to winners list for reporting
    winners.push({
      team: winner,
      score: winnerScore,
      fromRow: (score1 >= score2) ? i : (i + 1),
      toRow: round2Row
    });
    
    // Place winner in Round 2 column (column E)
    teamScoreSheet.getRange(round2Row, 5).setValue(winner);
    
    // Reset the score to 0 in Round 2 score column (column F)
    teamScoreSheet.getRange(round2Row, 6).setValue(0);
  }
  
  // Format the Round 2 section
  const round2Range = teamScoreSheet.getRange("E2:F9");
  round2Range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // Log the advancement operation
  console.log(`Advanced ${winners.length} teams to Round 2: ${winners.map(w => w.team).join(', ')}`);
  
  // Return winners information
  return winners;
}

/**
 * Updates Round 2 scores based on points earned since March 9
 */
function updateRoundTwoScores() {
  return updateRoundScores(2, 5, 6, ROUND_DATES.ROUND_2);
}

/**
 * Advances winners from Round 2 to Round 3 (semifinals)
 */
function advanceToRoundThree() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamScoreSheet = ss.getSheetByName("Team Scores");
  
  // Determine winners from Round 2
  const winners = [];
  
  // Process each pair of teams in Round 2
  for (let i = 2; i <= 8; i += 2) {
    const team1 = teamScoreSheet.getRange(i, 5).getValue();
    const team2 = teamScoreSheet.getRange(i + 1, 5).getValue();
    const score1 = teamScoreSheet.getRange(i, 6).getValue() || 0;
    const score2 = teamScoreSheet.getRange(i + 1, 6).getValue() || 0;
    
    // Determine winner based on higher score
    const winner = (score1 >= score2) ? team1 : team2;
    const winnerScore = Math.max(score1, score2);
    
    // Calculate position for winner in Round 3
    const round3Row = Math.floor((i - 2) / 2) + 2;
    
    // Add to winners list for reporting
    winners.push({
      team: winner,
      score: winnerScore,
      fromRow: (score1 >= score2) ? i : (i + 1),
      toRow: round3Row
    });
    
    // Place winner in Round 3 column (column I)
    teamScoreSheet.getRange(round3Row, 9).setValue(winner);
    
    // Reset the score to 0 in Round 3 score column (column J)
    teamScoreSheet.getRange(round3Row, 10).setValue(0);
  }
  
  // Format the Round 3 section
  const round3Range = teamScoreSheet.getRange("I2:J5");
  round3Range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // Log the advancement operation
  console.log(`Advanced ${winners.length} teams to Round 3: ${winners.map(w => w.team).join(', ')}`);
  
  // Return winners information
  return winners;
}

/**
 * Updates Round 3 scores based on points earned since March 16
 */
function updateRoundThreeScores() {
  return updateRoundScores(3, 9, 10, ROUND_DATES.ROUND_3);
}

/**
 * Advances winners from Round 3 to Round 4 (finals)
 */
function advanceToRoundFour() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamScoreSheet = ss.getSheetByName("Team Scores");
  
  // Determine winners from Round 3
  const winners = [];
  
  // Process each pair of teams in Round 3
  for (let i = 2; i <= 4; i += 2) {
    const team1 = teamScoreSheet.getRange(i, 9).getValue();
    const team2 = teamScoreSheet.getRange(i + 1, 9).getValue();
    const score1 = teamScoreSheet.getRange(i, 10).getValue() || 0;
    const score2 = teamScoreSheet.getRange(i + 1, 10).getValue() || 0;
    
    // Determine winner based on higher score
    const winner = (score1 >= score2) ? team1 : team2;
    const winnerScore = Math.max(score1, score2);
    
    // Calculate position for winner in Round 4
    const round4Row = Math.floor((i - 2) / 2) + 2;
    
    // Add to winners list for reporting
    winners.push({
      team: winner,
      score: winnerScore,
      fromRow: (score1 >= score2) ? i : (i + 1),
      toRow: round4Row
    });
    
    // Place winner in Round 4 column (column M)
    teamScoreSheet.getRange(round4Row, 13).setValue(winner);
    
    // Reset the score to 0 in Round 4 score column (column N)
    teamScoreSheet.getRange(round4Row, 14).setValue(0);
  }
  
  // Format the Round 4 section
  const round4Range = teamScoreSheet.getRange("M2:N3");
  round4Range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // Log the advancement operation
  console.log(`Advanced ${winners.length} teams to Finals: ${winners.map(w => w.team).join(', ')}`);
  
  // Return winners information
  return winners;
}

/**
 * Updates Round 4 scores based on points earned since March 23
 */
function updateRoundFourScores() {
  return updateRoundScores(4, 13, 14, ROUND_DATES.ROUND_4);
}

/**
 * Determines the champion based on Round 4 scores
 */
function determineChampion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamScoreSheet = ss.getSheetByName("Team Scores");
  
  // Get the two finalists
  const team1 = teamScoreSheet.getRange(2, 13).getValue();
  const team2 = teamScoreSheet.getRange(3, 13).getValue();
  const score1 = teamScoreSheet.getRange(2, 14).getValue() || 0;
  const score2 = teamScoreSheet.getRange(3, 14).getValue() || 0;
  
  // Determine champion based on higher score
  const champion = (score1 >= score2) ? team1 : team2;
  const championScore = Math.max(score1, score2);
  
  // Place champion in the champion column (column Q)
  teamScoreSheet.getRange(2, 17).setValue(champion);
  teamScoreSheet.getRange(2, 18).setValue(championScore);
  
  // Format the champion section with gold background
  const championRange = teamScoreSheet.getRange("Q2:R2");
  championRange.setBackground("#FFD700");
  championRange.setFontWeight("bold");
  championRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  // Log the champion
  console.log(`Champion determined: ${champion} with score ${championScore}`);
  
  return {
    champion: champion,
    score: championScore,
    runner_up: (score1 >= score2) ? team2 : team1,
    runner_up_score: Math.min(score1, score2)
  };
}

/**
 * Generic function to update scores for any round
 * @param {number} roundNumber - The round number (2, 3, or 4)
 * @param {number} teamColumn - The column index for team names
 * @param {number} scoreColumn - The column index for scores
 * @param {Date} startDate - The start date for this round
 */
function updateRoundScores(roundNumber, teamColumn, scoreColumn, startDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamScoreSheet = ss.getSheetByName("Team Scores");
  const teamsSheet = ss.getSheetByName("Teams");
  const auditLogSheet = ss.getSheetByName("Audit Log");
  
  // Get the round's team count and row range
  const teamCounts = { 2: 8, 3: 4, 4: 2 };
  const teamCount = teamCounts[roundNumber] || 8;
  
  // Get the teams for this round
  const roundTeams = [];
  for (let i = 2; i <= teamCount + 1; i++) {
    const teamName = teamScoreSheet.getRange(i, teamColumn).getValue();
    if (teamName) roundTeams.push(teamName);
  }
  
  console.log(`Found ${roundTeams.length} teams in Round ${roundNumber}: ${roundTeams.join(', ')}`);
  
  // Get all audit log entries
  const auditData = auditLogSheet.getDataRange().getValues();
  
  // Create mapping of members to teams
  const memberTeamMap = new Map();
  const teamData = teamsSheet.getDataRange().getValues();
  
  for (let i = 1; i < teamData.length; i++) {
    const teamName = teamData[i][0];
    const memberMnemonic = String(teamData[i][1]).toLowerCase();
    
    if (teamName && memberMnemonic) {
      memberTeamMap.set(memberMnemonic, teamName);
    }
  }
  
  // Calculate scores for each team
  const teamScores = {};
  roundTeams.forEach(team => { teamScores[team] = 0; });
  
  // Process audit log entries
  for (let i = 1; i < auditData.length; i++) {
    const timestamp = auditData[i][0]; // Timestamp in column A
    if (!(timestamp instanceof Date) || timestamp < startDate) continue;
    
    const mnemonic = String(auditData[i][1]).toLowerCase(); // Mnemonic in column B
    const earnedPoints = Number(auditData[i][8]) || 0; // Points in column I
    
    // Get the team for this member
    const teamName = memberTeamMap.get(mnemonic);
    
    // If member is in a team for this round, add points
    if (teamName && roundTeams.includes(teamName)) {
      teamScores[teamName] += earnedPoints;
    }
  }
  
  // Update scores in the sheet
  for (let i = 0; i < roundTeams.length; i++) {
    const teamName = roundTeams[i];
    const row = i + 2; // Start at row 2
    const score = teamScores[teamName] || 0;
    
    // Update the score column
    teamScoreSheet.getRange(row, scoreColumn).setValue(score);
    console.log(`Updated ${teamName} with score ${score} in Round ${roundNumber}`);
  }
  
  return `Updated scores for ${roundTeams.length} teams in Round ${roundNumber}.`;
}
