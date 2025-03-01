/**
 * Show dialog for configuring automatic processing triggers
 */
function showTriggerSetupDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
      .form-group { margin-bottom: 15px; }
      select, input { width: 100%; padding: 5px; margin-top: 3px; }
      label { font-weight: bold; }
      button { background-color: #4285f4; color: white; border: none; padding: 8px 12px; cursor: pointer; }
      .header { font-size: 16px; font-weight: bold; margin-bottom: 15px; }
      .footer { margin-top: 20px; text-align: right; }
    </style>
    <div class="header">Configure Automatic Processing</div>
    <form id="triggerForm">
      <div class="form-group">
        <label for="syncResponses">Sync Responses Frequency:</label>
        <select id="syncResponses">
          <option value="1">Every 1 minute</option>
          <option value="3" selected>Every 3 minutes</option>
          <option value="5">Every 5 minutes</option>
          <option value="10">Every 10 minutes</option>
          <option value="15">Every 15 minutes</option>
          <option value="30">Every 30 minutes</option>
          <option value="60">Every hour</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="processQueue">Process Queue Frequency:</label>
        <select id="processQueue">
          <option value="1">Every 1 minute</option>
          <option value="3" selected>Every 3 minutes</option>
          <option value="5">Every 5 minutes</option>
          <option value="10">Every 10 minutes</option>
          <option value="15">Every 15 minutes</option>
          <option value="30">Every 30 minutes</option>
          <option value="60">Every hour</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="updateLeaderboard">Update Leaderboard Frequency:</label>
        <select id="updateLeaderboard">
          <option value="5">Every 5 minutes</option>
          <option value="10">Every 10 minutes</option>
          <option value="15" selected>Every 15 minutes</option>
          <option value="30">Every 30 minutes</option>
          <option value="60">Every hour</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="archiveData">Archive Old Data Frequency:</label>
        <select id="archiveData">
          <option value="6">Every 6 hours</option>
          <option value="12" selected>Every 12 hours</option>
          <option value="24">Every 24 hours</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="updateQuestionsHour">Update Daily Questions at Hour:</label>
        <select id="updateQuestionsHour">
          ${Array.from({length: 24}, (_, i) => 
            `<option value="${i}" ${i === 0 ? 'selected' : ''}>${i}:00</option>`
          ).join('')}
        </select>
      </div>
      
      <div class="form-group">
        <label for="weeklyWinnersDay">Weekly Winners Day:</label>
        <select id="weeklyWinnersDay">
          <option value="1">Monday</option>
          <option value="2">Tuesday</option>
          <option value="3">Wednesday</option>
          <option value="4">Thursday</option>
          <option value="5">Friday</option>
          <option value="6">Saturday</option>
          <option value="7" selected>Sunday</option>
        </select>
      </div>
      
      <div class="form-group">
        <label for="weeklyWinnersHour">Weekly Winners Hour:</label>
        <select id="weeklyWinnersHour">
          ${Array.from({length: 24}, (_, i) => 
            `<option value="${i}" ${i === 23 ? 'selected' : ''}>${i}:00</option>`
          ).join('')}
        </select>
      </div>
      
      <div class="form-group">
        <label for="enableFormTrigger">Enable Form Submission Trigger:</label>
        <select id="enableFormTrigger">
          <option value="yes" selected>Yes</option>
          <option value="no">No</option>
        </select>
      </div>
      
      <div class="footer">
        <button type="button" onclick="saveTriggers()">Save & Setup Triggers</button>
      </div>
    </form>
    
    <script>
      function saveTriggers() {
        const config = {
          syncResponses: document.getElementById('syncResponses').value,
          processQueue: document.getElementById('processQueue').value,
          updateLeaderboard: document.getElementById('updateLeaderboard').value,
          archiveData: document.getElementById('archiveData').value,
          updateQuestionsHour: document.getElementById('updateQuestionsHour').value,
          weeklyWinnersDay: document.getElementById('weeklyWinnersDay').value,
          weeklyWinnersHour: document.getElementById('weeklyWinnersHour').value,
          enableFormTrigger: document.getElementById('enableFormTrigger').value === 'yes'
        };
        
        google.script.run
          .withSuccessHandler(function() {
            alert('Triggers have been set up successfully!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Error setting up triggers: ' + error);
          })
          .setupTriggersWithConfig(config);
      }
    </script>
  `)
    .setWidth(400)
    .setHeight(550)
    .setTitle('Setup Automatic Processing');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Setup Automatic Processing');
}

/**
 * Set up triggers based on user configuration
 */
function setupTriggersWithConfig(config) {
  // Convert string values to numbers
  const settings = {
    syncResponses: parseInt(config.syncResponses),
    processQueue: parseInt(config.processQueue),
    updateLeaderboard: parseInt(config.updateLeaderboard),
    archiveData: parseInt(config.archiveData),
    updateQuestionsHour: parseInt(config.updateQuestionsHour),
    weeklyWinnersDay: parseInt(config.weeklyWinnersDay),
    weeklyWinnersHour: parseInt(config.weeklyWinnersHour),
    enableFormTrigger: config.enableFormTrigger
  };
  
  // Save configuration to script properties
  PropertiesService.getScriptProperties().setProperty('triggerConfig', JSON.stringify(settings));
  
  // Clear existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // 1. Create trigger for syncResponses
  ScriptApp.newTrigger('syncResponses')
    .timeBased()
    .everyMinutes(settings.syncResponses)
    .create();
  
  // 2. Create trigger for processQueue
  ScriptApp.newTrigger('processQueue')
    .timeBased()
    .everyMinutes(settings.processQueue)
    .create();
  
  // 3. Create trigger for updateLeaderboard
  ScriptApp.newTrigger('updateLeaderboard')
    .timeBased()
    .everyMinutes(settings.updateLeaderboard)
    .create();
  
  // 4. Create trigger for archiveOldData
  ScriptApp.newTrigger('archiveOldData')
    .timeBased()
    .everyHours(settings.archiveData)
    .create();
  
  // 5. Create trigger for updateDailyQuestions
  ScriptApp.newTrigger('updateDailyQuestions')
    .timeBased()
    .atHour(settings.updateQuestionsHour)
    .everyDays(1)
    .create();
  
  // 6. Create trigger for determineWeeklyWinners
  const weekDay = getWeekDayEnum(settings.weeklyWinnersDay);
  ScriptApp.newTrigger('determineWeeklyWinners')
    .timeBased()
    .onWeekDay(weekDay)
    .atHour(settings.weeklyWinnersHour)
    .create();
  
  // 7. Create form submission trigger if enabled
  if (settings.enableFormTrigger) {
    try {
      const form = FormApp.openById(FORM_ID);
      ScriptApp.newTrigger('onFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
      console.log("✅ Form trigger created successfully");
    } catch (e) {
      console.error("❌ Error creating form trigger: " + e.message);
      throw new Error("Could not set up form trigger. Please check your form ID.");
    }
  }
  
  // 8. Create trigger for clearing caches every 6 hours
  ScriptApp.newTrigger('clearCaches')
    .timeBased()
    .everyHours(6)
    .create();
  
  console.log("✅ All triggers set up successfully with user configuration");
  return true;
}

/**
 * Helper function to convert day number to ScriptApp.WeekDay enum
 */
function getWeekDayEnum(dayNumber) {
  switch (parseInt(dayNumber)) {
    case 1: return ScriptApp.WeekDay.MONDAY;
    case 2: return ScriptApp.WeekDay.TUESDAY;
    case 3: return ScriptApp.WeekDay.WEDNESDAY;
    case 4: return ScriptApp.WeekDay.THURSDAY;
    case 5: return ScriptApp.WeekDay.FRIDAY;
    case 6: return ScriptApp.WeekDay.SATURDAY;
    case 7: return ScriptApp.WeekDay.SUNDAY;
    default: return ScriptApp.WeekDay.SUNDAY;
  }
}

function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  console.log("=== CURRENT TRIGGERS ===");
  if (triggers.length === 0) {
    console.log("No triggers found! You need to run setupTriggers().");
  } else {
    triggers.forEach((trigger, index) => {
      console.log(`[${index + 1}] Function: ${trigger.getHandlerFunction()}`);
      console.log(`    Type: ${getTriggerTypeString(trigger.getEventType())}`);
      
      // For time-based triggers, show frequency
      if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        const frequency = getTriggerFrequency(trigger);
        console.log(`    Frequency: ${frequency}`);
      }
    });
  }
  
  // Check if gradeResponses is directly or indirectly triggered
  const hasGradeResponsesTrigger = triggers.some(t => t.getHandlerFunction() === 'gradeResponses');
  const hasProcessQueueTrigger = triggers.some(t => t.getHandlerFunction() === 'processQueue');
  
  console.log("\n=== GRADING WORKFLOW STATUS ===");
  if (!hasGradeResponsesTrigger) {
    console.log("❌ No direct trigger for gradeResponses() found");
    if (hasProcessQueueTrigger) {
      console.log("✓ ProcessQueue trigger found - grading should happen through the queue");
    } else {
      console.log("❌ No processQueue trigger found either - grading workflow is broken");
    }
  } else {
    console.log("✓ Direct trigger for gradeResponses() exists");
  }
  
  // Check form trigger
  const hasFormTrigger = triggers.some(t => 
    t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT
  );
  console.log("\n=== FORM SUBMISSION HANDLING ===");
  if (hasFormTrigger) {
    console.log("✓ Form submission trigger exists");
  } else {
    console.log("❌ No form submission trigger found");
    console.log("   Form submissions won't be processed automatically");
  }
  
  // Check if we have a direct trigger for syncResponses
  const hasSyncTrigger = triggers.some(t => t.getHandlerFunction() === 'syncResponses');
  console.log("\n=== RESPONSE SYNCING STATUS ===");
  if (hasSyncTrigger) {
    console.log("✓ Direct trigger for syncResponses() exists");
  } else {
    console.log("❌ No syncResponses trigger found");
    console.log("   Form responses won't be synced automatically");
  }
}

// Helper function to get readable trigger type
function getTriggerTypeString(eventType) {
  switch(eventType) {
    case ScriptApp.EventType.CLOCK:
      return "Time-based";
    case ScriptApp.EventType.ON_EDIT:
      return "On Edit";
    case ScriptApp.EventType.ON_FORM_SUBMIT:
      return "On Form Submit";
    case ScriptApp.EventType.ON_OPEN:
      return "On Open";
    default:
      return "Unknown";
  }
}

// Helper function to identify trigger frequency
function getTriggerFrequency(trigger) {
  // We can't directly access the frequency, but we can check the trigger source
  const triggerSource = trigger.getTriggerSourceId();
  
  // Try to make an educated guess based on creation time
  const now = new Date();
  const handlerFunction = trigger.getHandlerFunction();
  
  return "Time-based trigger (specific frequency not accessible)";
}
