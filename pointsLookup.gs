// Code.gs file
function doGet() {
  return HtmlService.createTemplateFromFile('PointsLookup')
      .evaluate()
      .setTitle('March Madness Points Lookup')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPointsForMnemonic(mnemonic) {
  // Trim and lowercase for consistent comparison
  mnemonic = (mnemonic || "").trim().toLowerCase();
  
  if (!mnemonic) {
    return { success: false, message: "Please enter your mnemonic" };
  }
  
  try {
    // Open the spreadsheet - replace with your spreadsheet ID
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Individual Leaderboard'); // Replace with your actual sheet name
    
    // Get all data at once to minimize API calls
    const data = sheet.getDataRange().getValues();
    
    // Get header row to find column indexes
    const headers = data[0].map(h => String(h).toLowerCase());
    const mnemonicIdx = headers.findIndex(h => h.includes('mnemonic'));
    const scoreIdx = headers.findIndex(h => h.includes('score') || h.includes('total') || h.includes('points'));
    const nameIdx = headers.findIndex(h => h.includes('name'));
    const roleIdx = headers.findIndex(h => h.includes('role'));
    
    // Check if required columns exist
    if (mnemonicIdx === -1 || scoreIdx === -1) {
      return { success: false, message: "Required columns not found in spreadsheet" };
    }
    
    // Look for matching mnemonic
    for (let i = 1; i < data.length; i++) {
      const rowMnemonic = String(data[i][mnemonicIdx] || "").trim().toLowerCase();
      
      if (rowMnemonic === mnemonic) {
        return { 
          success: true,
          score: data[i][scoreIdx] || 0,
          name: nameIdx >= 0 ? data[i][nameIdx] || "" : "",
          role: roleIdx >= 0 ? data[i][roleIdx] || "" : ""
        };
      }
    }
    
    // If no match found
    return { success: false, message: "Mnemonic not found. Please check your entry." };
    
  } catch (error) {
    return { success: false, message: "Error looking up points: " + error.message };
  }
}