/**
 * Question Bank Editor - Smart Comma Separation Version
 * This file contains all functions related to the Question Bank Editor feature
 */

/**
 * Show the Question Bank Editor dialog
 */
function showQuestionBankEditor() {
  const html = HtmlService.createHtmlOutput(getQuestionBankEditorHtml())
    .setWidth(700)
    .setHeight(600)
    .setTitle('Question Bank Editor');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Question Bank Editor');
}

/**
 * Show edit dialog for an existing question
 */
function editExistingQuestion() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask for question ID to edit
  const response = ui.prompt(
    'Edit Existing Question',
    'Enter the Question ID to edit (e.g., Q001):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const questionId = response.getResponseText().trim().toUpperCase();
  
  // Find the question in the Question Bank
  const questionData = findQuestionById(questionId);
  
  if (!questionData) {
    ui.alert('Error', `Question ID ${questionId} not found in the Question Bank.`, ui.ButtonSet.OK);
    return;
  }
  
  // Show the editor with pre-filled data
  const html = HtmlService.createHtmlOutput(getQuestionBankEditorHtml(questionData))
    .setWidth(700)
    .setHeight(600)
    .setTitle(`Edit Question ${questionId}`);
  
  SpreadsheetApp.getUi().showModalDialog(html, `Edit Question ${questionId}`);
}

/**
 * Find a question by ID in the Question Bank
 */
function findQuestionById(questionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.QUESTION_BANK);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === questionId) {
      // Retrieve the correct answer - could be comma separated
      let correctAnswer = data[i][9] || '';
      
      return {
        rowIndex: i + 1,
        date: data[i][0],
        questionId: data[i][1],
        question: data[i][2],
        options: [
          data[i][3], // Option A
          data[i][4], // Option B
          data[i][5], // Option C
          data[i][6], // Option D
          data[i][7], // Option E
          data[i][8]  // Option F
        ],
        correctAnswer: correctAnswer,
        type: data[i][10],
        targetRole: data[i][11],
        points: data[i][12],
        imageUrl: data[i][13] || ''
      };
    }
  }
  
  return null;
}


/**
 * Process the form submission and add/update the question in the Question Bank
 * UPDATED: Fixed timezone issues and improved answer formatting
 */
function processQuestionForm(formData) {
  try {
    console.log("Processing form data:", JSON.stringify(formData));
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.QUESTION_BANK);
    
    // Format date correctly - FIX: Set to noon to avoid timezone issues
    let questionDate;
    if (formData.date) {
      questionDate = new Date(formData.date);
      // Set to noon local time to avoid date rollover due to timezone
      questionDate.setHours(12, 0, 0, 0);
    } else {
      questionDate = new Date();
      questionDate.setHours(12, 0, 0, 0);
    }
    
    console.log("Date selected:", formData.date);
    console.log("Date to be stored:", questionDate);
    
    // Process correct answers
    let correctAnswer = '';
    if (formData.type === 'Multiple Select') {
      // For multiple select, join selected options with comma delimiter for smart processing
      const selectedOptions = [];
      for (let i = 0; i < formData.options.length; i++) {
        if (formData.correctOptions[i] && formData.options[i] && formData.options[i] !== 'N/A') {
          selectedOptions.push(formData.options[i]);
        }
      }
      correctAnswer = selectedOptions.join(',');
      console.log("Multiple select correct answers:", selectedOptions);
    } else if (formData.type === 'Multiple Choice') {
      // For multiple choice, get the single selected option
      const selectedIndex = formData.correctOptions.indexOf(true);
      if (selectedIndex >= 0 && selectedIndex < formData.options.length) {
        correctAnswer = formData.options[selectedIndex];
      }
    } else if (formData.type === 'Short Answer') {
      correctAnswer = formData.shortAnswerCorrect;
    }
    
    console.log("Final correct answer format:", correctAnswer);
    
    // Handle existing question (update)
    if (formData.isEdit && formData.rowIndex) {
      const rowIndex = parseInt(formData.rowIndex);
      
      // Update all fields in the row
      sheet.getRange(rowIndex, 1).setValue(questionDate);
      sheet.getRange(rowIndex, 2).setValue(formData.questionId);
      sheet.getRange(rowIndex, 3).setValue(formData.question);
      
      // Update options (columns 4-9)
      for (let i = 0; i < 6; i++) {
        const optionValue = i < formData.options.length ? formData.options[i] : '';
        sheet.getRange(rowIndex, 4 + i).setValue(optionValue === 'N/A' ? '' : optionValue);
      }
      
      // Update remaining fields
      sheet.getRange(rowIndex, 10).setValue(correctAnswer);
      sheet.getRange(rowIndex, 11).setValue(formData.type);
      sheet.getRange(rowIndex, 12).setValue(formData.targetRole);
      sheet.getRange(rowIndex, 13).setValue(parseInt(formData.points) || 1);
      sheet.getRange(rowIndex, 14).setValue(formData.imageUrl || '');
      
      return {
        success: true,
        message: `Question ${formData.questionId} updated successfully!`,
        questionId: formData.questionId
      };
    } else {
      // New question (append)
      const newRow = [
        questionDate,
        formData.questionId,
        formData.question
      ];
      
      // Add options
      for (let i = 0; i < 6; i++) {
        const optionValue = i < formData.options.length ? formData.options[i] : '';
        newRow.push(optionValue === 'N/A' ? '' : optionValue);
      }
      
      // Add remaining fields
      newRow.push(
        correctAnswer,           // Correct Answer
        formData.type,           // Question Type
        formData.targetRole,     // Target Role
        parseInt(formData.points) || 1,  // Points
        formData.imageUrl || ''  // Image URL
      );
      
      // Append to sheet
      sheet.appendRow(newRow);
      
      return {
        success: true,
        message: `Question ${formData.questionId} added successfully!`,
        questionId: formData.questionId
      };
    }
  } catch (error) {
    console.error("Error processing question form:", error.message, error.stack);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Get the next available question ID
 */
function getNextQuestionId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.QUESTION_BANK);
  const data = sheet.getDataRange().getValues();
  
  let highestNumber = 0;
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const questionId = data[i][1];
    if (questionId && questionId.startsWith('Q')) {
      const num = parseInt(questionId.substring(1));
      if (!isNaN(num) && num > highestNumber) {
        highestNumber = num;
      }
    }
  }
  
  const nextNumber = highestNumber + 1;
  return `Q${nextNumber.toString().padStart(3, '0')}`;
}

/**
 * Generate the HTML for the Question Bank Editor - Fixed version
 */
function getQuestionBankEditorHtml(questionData = null) {
  // Get next question ID if this is a new question
  const nextQuestionId = questionData ? questionData.questionId : getNextQuestionId();
  
  // Determine if this is an edit or new question
  const isEdit = !!questionData;
  
  // Generate helper text for multiple select
  const multipleSelectHelp = "For multiple select questions, choose all correct answers. The system will handle comma separation intelligently.";
  
  // Debug the question data if available
  if (isEdit && questionData) {
    console.log("Edit mode with question data:", JSON.stringify({
      id: questionData.questionId,
      type: questionData.type,
      options: questionData.options,
      correctAnswer: questionData.correctAnswer
    }));
  }
  
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 16px;
      font-size: 14px;
    }
    .form-group {
      margin-bottom: 12px;
    }
    label {
      display: block;
      margin-bottom: 4px;
      font-weight: bold;
    }
    input[type="text"], input[type="date"], input[type="number"], select, textarea {
      width: 100%;
      padding: 8px;
      border: 1px solid #ccc;
      border-radius: 4px;
      box-sizing: border-box;
    }
    textarea {
      min-height: 80px;
    }
    .option-row {
      display: flex;
      align-items: center;
      margin-bottom: 8px;
    }
    .option-row input[type="checkbox"] {
      margin-right: 8px;
    }
    .option-row input[type="text"] {
      flex-grow: 1;
    }
    .options-container {
      margin-top: 8px;
      margin-bottom: 16px;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
    }
    .button-container {
      margin-top: 16px;
      text-align: right;
    }
    button {
      padding: 8px 16px;
      background-color: #4285f4;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
    button:hover {
      background-color: #3b78e7;
    }
    .help-text {
      font-size: 12px;
      color: #666;
      margin-top: 2px;
    }
    .warning {
      background-color: #fff3cd;
      border: 1px solid #ffeeba;
      color: #856404;
      padding: 8px;
      margin-bottom: 16px;
      border-radius: 4px;
    }
    .debug-info {
      font-family: monospace;
      background-color: #f5f5f5;
      padding: 8px;
      margin-top: 20px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 12px;
      display: none;
    }
  </style>
</head>
<body>
  <h2>${isEdit ? 'Edit Question' : 'Add New Question'}</h2>
  
  <form id="questionForm">
    <!-- Hidden fields for edit mode -->
    <input type="hidden" id="isEdit" name="isEdit" value="${isEdit ? 'true' : 'false'}">
    ${isEdit ? `<input type="hidden" id="rowIndex" name="rowIndex" value="${questionData.rowIndex}">` : ''}
    
    <div class="form-group">
      <label for="questionId">Question ID:</label>
      <input type="text" id="questionId" name="questionId" value="${nextQuestionId}" ${isEdit ? 'readonly' : ''} required>
      <div class="help-text">Format: Q001, Q002, etc. Will be auto-generated for new questions.</div>
    </div>
    
    <div class="form-group">
      <label for="date">Date:</label>
      <input type="date" id="date" name="date" value="${isEdit && questionData.date ? new Date(questionData.date).toISOString().split('T')[0] : new Date().toISOString().split('T')[0]}" required>
      <div class="help-text">The date this question should appear in the form.</div>
    </div>
    
    <div class="form-group">
      <label for="type">Question Type:</label>
      <select id="type" name="type" onchange="updateFormForType()" required>
        <option value="Multiple Choice" ${isEdit && questionData.type === 'Multiple Choice' ? 'selected' : ''}>Multiple Choice (single answer)</option>
        <option value="Multiple Select" ${isEdit && questionData.type === 'Multiple Select' ? 'selected' : ''}>Multiple Select (multiple answers)</option>
        <option value="Short Answer" ${isEdit && questionData.type === 'Short Answer' ? 'selected' : ''}>Short Answer</option>
      </select>
    </div>
    
    <div class="form-group">
      <label for="targetRole">Target Role:</label>
      <select id="targetRole" name="targetRole" required>
        <option value="RN" ${isEdit && questionData.targetRole === 'RN' ? 'selected' : ''}>RN</option>
        <option value="PCA" ${isEdit && questionData.targetRole === 'PCA' ? 'selected' : ''}>PCA</option>
      </select>
      <div class="help-text">Which role should see this question.</div>
    </div>
    
    <div class="form-group">
      <label for="question">Question Text:</label>
      <textarea id="question" name="question" required>${isEdit ? questionData.question : ''}</textarea>
    </div>
    
    <div id="optionsContainer" class="options-container">
      <label>Answer Options:</label>
      <div class="help-text">Check the box for correct answer(s).</div>
      
      <div id="multipleSelectWarning" class="warning" style="display: none;">
        ${multipleSelectHelp}
      </div>
      
      <div id="optionsContent">
        <!-- Options will be dynamically added here -->
      </div>
    </div>
    
    <div id="shortAnswerContainer" class="form-group" style="display: none;">
      <label for="shortAnswerCorrect">Correct Answer:</label>
      <input type="text" id="shortAnswerCorrect" name="shortAnswerCorrect" value="${isEdit && questionData.type === 'Short Answer' ? questionData.correctAnswer : ''}">
      <div class="help-text">For short answer questions, enter the correct answer text.</div>
    </div>
    
    <div class="form-group">
      <label for="points">Points:</label>
      <input type="number" id="points" name="points" min="1" value="${isEdit ? questionData.points : '1'}" required>
    </div>
    
    <div class="form-group">
      <label for="imageUrl">Image URL (optional):</label>
      <input type="text" id="imageUrl" name="imageUrl" value="${isEdit && questionData.imageUrl ? questionData.imageUrl : ''}">
      <div class="help-text">If the question includes an image, enter its URL here.</div>
    </div>
    
    <!-- Debug info - hidden by default -->
    <div id="debugInfo" class="debug-info"></div>
    
    <div class="button-container">
      <button type="button" onclick="submitForm()">${isEdit ? 'Update' : 'Add'} Question</button>
    </div>
  </form>

  <script>
    // Debug function
    function debugLog(message, data) {
      console.log(message, data);
      const debugEl = document.getElementById('debugInfo');
      if (debugEl) {
        debugEl.innerHTML += message + ': ' + JSON.stringify(data) + '<br>';
      }
    }
    
    // On page load
    document.addEventListener('DOMContentLoaded', function() {
      updateFormForType();
      
      // Initialize the form with existing data in edit mode
      if (document.getElementById('isEdit').value === 'true') {
        debugLog('Edit mode initialized', {
          type: document.getElementById('type').value,
          correctAnswer: ${isEdit ? JSON.stringify(questionData.correctAnswer || '') : '""'}
        });
      }
    });
    
    // Update form based on question type
    function updateFormForType() {
      const type = document.getElementById('type').value;
      const optionsContainer = document.getElementById('optionsContainer');
      const shortAnswerContainer = document.getElementById('shortAnswerContainer');
      const multipleSelectWarning = document.getElementById('multipleSelectWarning');
      
      debugLog('Updating form for type', type);
      
      if (type === 'Short Answer') {
        optionsContainer.style.display = 'none';
        shortAnswerContainer.style.display = 'block';
        multipleSelectWarning.style.display = 'none';
      } else {
        optionsContainer.style.display = 'block';
        shortAnswerContainer.style.display = 'none';
        
        // Show warning for multiple select
        if (type === 'Multiple Select') {
          multipleSelectWarning.style.display = 'block';
        } else {
          multipleSelectWarning.style.display = 'none';
        }
        
        // Generate option fields
        generateOptionFields(type);
      }
    }
    
    // Generate option fields based on type
    function generateOptionFields(type) {
      const optionsContent = document.getElementById('optionsContent');
      optionsContent.innerHTML = '';
      
      const isMultipleSelect = type === 'Multiple Select';
      const inputType = isMultipleSelect ? 'checkbox' : 'radio';
      const inputName = isMultipleSelect ? 'correctOptions[]' : 'correctOption';
      
      // Get existing values if editing
      const isEdit = document.getElementById('isEdit').value === 'true';
      let existingOptions = [];
      let correctAnswers = [];
      
      // FIXED SECTION: Properly handle correct answers for different question types
      if (isEdit) {
        ${isEdit ? `
          // Get the non-empty options
          existingOptions = ${JSON.stringify(questionData.options.filter(Boolean))};
          
          // Get the correct answers
          const correctAnswerStr = ${JSON.stringify(questionData.correctAnswer || '')};
          debugLog('Raw correctAnswer', correctAnswerStr);
          
          if (isMultipleSelect) {
            // For multiple select, split by comma and trim
            correctAnswers = correctAnswerStr.split(',').map(a => a.trim());
          } else {
            // For single select, just use as-is
            correctAnswers = [correctAnswerStr.trim()];
          }
          
          debugLog('Parsed correctAnswers', correctAnswers);
          debugLog('Existing options', existingOptions);
        ` : ''}
      }
      
      // Create 6 option fields (or fewer if editing with fewer options)
      const numOptions = isEdit ? Math.max(existingOptions.length, 6) : 6;
      
      for (let i = 0; i < numOptions; i++) {
        const optionValue = isEdit && i < existingOptions.length ? existingOptions[i] : '';
        
        // FIXED: Better matching for correct answers
        let isCorrect = false;
        if (isEdit && optionValue) {
          if (isMultipleSelect) {
            // For multiple select, check if any correct answer matches this option
            isCorrect = correctAnswers.some(answer => {
              return answer.toLowerCase() === optionValue.toLowerCase();
            });
          } else {
            // For single select, just check the first (and only) correct answer
            isCorrect = correctAnswers.length > 0 && 
                       correctAnswers[0].toLowerCase() === optionValue.toLowerCase();
          }
        }
        
        const row = document.createElement('div');
        row.className = 'option-row';
        
        const checkbox = document.createElement('input');
        checkbox.type = inputType;
        checkbox.name = inputName;
        checkbox.value = i;
        checkbox.checked = isCorrect;
        checkbox.id = 'option_' + i;
        
        const label = document.createElement('label');
        label.setAttribute('for', 'option_' + i);
        label.style.marginRight = '8px';
        label.textContent = String.fromCharCode(65 + i) + ':'; // A, B, C, etc.
        
        const textInput = document.createElement('input');
        textInput.type = 'text';
        textInput.name = 'options[]';
        textInput.placeholder = 'Option ' + String.fromCharCode(65 + i) + ' (or enter N/A if not used)';
        textInput.value = optionValue;
        
        row.appendChild(checkbox);
        row.appendChild(label);
        row.appendChild(textInput);
        optionsContent.appendChild(row);
      }
    }
    
    // Submit the form
    function submitForm() {
      const form = document.getElementById('questionForm');
      const formData = {
        isEdit: form.isEdit.value === 'true',
        rowIndex: form.isEdit.value === 'true' ? form.rowIndex.value : null,
        questionId: form.questionId.value,
        date: form.date.value,
        type: form.type.value,
        targetRole: form.targetRole.value,
        question: form.question.value,
        points: form.points.value,
        imageUrl: form.imageUrl.value,
        options: [],
        correctOptions: []
      };
      
      // Get options and correct selections
      if (formData.type !== 'Short Answer') {
        const optionInputs = document.querySelectorAll('input[name="options[]"]');
        const correctInputs = document.querySelectorAll('input[name="correctOptions[]"], input[name="correctOption"]');
        
        optionInputs.forEach((input, index) => {
          if (input.value && input.value !== 'N/A') {
            formData.options.push(input.value);
            formData.correctOptions[index] = correctInputs[index].checked;
          }
        });
        
        debugLog('Form submission - options', formData.options);
        debugLog('Form submission - correctOptions', formData.correctOptions);
      } else {
        formData.shortAnswerCorrect = document.getElementById('shortAnswerCorrect').value;
      }
      
      // Validate form
      if (!formData.questionId || !formData.question) {
        alert('Please fill in all required fields.');
        return;
      }
      
      if (formData.type !== 'Short Answer' && formData.options.length === 0) {
        alert('Please add at least one valid answer option.');
        return;
      }
      
      if (formData.type !== 'Short Answer') {
        const hasCorrectAnswer = formData.correctOptions.some(isCorrect => isCorrect);
        if (!hasCorrectAnswer) {
          alert('Please select at least one correct answer.');
          return;
        }
      } else if (!formData.shortAnswerCorrect) {
        alert('Please enter the correct answer for the short answer question.');
        return;
      }
      
      // Submit to server
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            alert(result.message);
            google.script.host.close();
          } else {
            alert('Error: ' + result.message);
          }
        })
        .withFailureHandler(function(error) {
          alert('Error: ' + error.message);
        })
        .processQuestionForm(formData);
    }
  </script>
</body>
</html>
  `;
}

/**
 * This function should be called from your existing onOpen function 
 * to integrate the Question Bank Editor with your menu
 */
function addQuestionBankEditorToMenu() {
  // This function should be called from your existing onOpen function
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Quiz Admin');
  
  // Add your existing menu items here
  // Example:
  menu
    .addItem('Process Pending Responses', 'processQueue')
    .addItem('Flush All Pending Responses', 'flushQueue')
    .addItem('Grade Responses', 'gradeResponses')
    .addItem('Sync Responses', 'syncResponses')
    .addSeparator()
    
    // Question Bank Editor items
    .addItem('Add New Question', 'showQuestionBankEditor')
    .addItem('Edit Existing Question', 'editExistingQuestion')
    .addSeparator()
    
    // Add the rest of your menu items
    .addItem('Update Leaderboard', 'updateLeaderboard')
    // ... etc.
    
    .addToUi();
}

// Helper function to update your onOpen
function updateOnOpenForEditor() {
  // Example code showing how to integrate with your onOpen function
  
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Quiz Admin');
  
  // Add all your existing menu items
  
  // Add question bank editor items
  menu.addSeparator()
    .addItem('Add New Question', 'showQuestionBankEditor')
    .addItem('Edit Existing Question', 'editExistingQuestion');
  
  // Continue with the rest of your menu items
  
  menu.addToUi();
}

/**
 * Smart function to handle comma-separated options for multiple select questions
 * This should work with your existing code for handling answers
 */
function processMultipleSelectAnswer(userAnswer, questionId, allOptions) {
  // For multiple select questions, we need to handle commas in the text
  if (!userAnswer) return [];
  
  // If we don't have the options, fall back to simple comma splitting
  if (!allOptions || allOptions.length === 0) {
    return userAnswer.split(',').map(item => item.trim());
  }
  
  // Smart parsing using the known options
  let selectedOptions = [];
  let remainingAnswer = userAnswer;
  
  // For each possible option, check if it appears in the answer
  allOptions.forEach(option => {
    if (remainingAnswer.includes(option)) {
      selectedOptions.push(option);
      // Remove this option from the remaining string
      remainingAnswer = remainingAnswer.replace(option, '').replace(/^,\s*/, '');
    }
  });
  
  return selectedOptions;
}