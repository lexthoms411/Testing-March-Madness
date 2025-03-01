/**
 * Updates questions for today
 */
function updateDailyQuestions() {
  try {
    const form = FormApp.openById(FORM_ID);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const questionBank = ss.getSheetByName(SHEETS.QUESTION_BANK);
    
    if (!form || !questionBank) {
      throw new Error('Could not access form or question bank');
    }

    const sections = findFormSections(form);
    if (!sections.rn || !sections.pca) {
      throw new Error('Could not find RN or PCA sections');
    }

    clearQuestionsInSections(form, sections);
    Utilities.sleep(1000);

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
    addTodaysQuestions(form, questionBank, sections, today);

    logToSheet('Update Questions', 'SUCCESS', 'Daily questions updated successfully');
  } catch (error) {
    logToSheet('Update Questions', 'ERROR', error.message);
    throw error;
  }
}

function findFormSections(form) {
  const items = form.getItems(FormApp.ItemType.PAGE_BREAK);
  console.log('Finding form sections...');
  
  const sections = {
    rn: items.find(item => item.getTitle() === SECTIONS.RN) || null,
    pca: items.find(item => item.getTitle() === SECTIONS.PCA) || null
  };
  
  console.log('Found RN section:', sections.rn ? 'Yes' : 'No');
  console.log('Found PCA section:', sections.pca ? 'Yes' : 'No');
  
  return sections;
}

function clearQuestionsInSections(form, sections) {
  console.log('Starting to clear questions from sections...');
  const items = form.getItems();
  const toDelete = [];
  let inSection = false;
  let currentSection = '';

  items.forEach(item => {
    if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
      const title = item.getTitle();
      console.log('Found section:', title);
      inSection = (title === SECTIONS.RN || title === SECTIONS.PCA);
      currentSection = title;
    } else if (inSection) {
      console.log(`Marking item for deletion in ${currentSection}:`, item.getTitle());
      toDelete.push(item);
    }
  });

  console.log(`Total items to delete: ${toDelete.length}`);
  
  for (let i = toDelete.length - 1; i >= 0; i--) {
    form.deleteItem(toDelete[i].getIndex());
    Utilities.sleep(100);
  }
  
  console.log('Finished clearing old questions');
}

function addTodaysQuestions(form, questionBank, sections, today) {
  const data = questionBank.getDataRange().getValues();
  let questionsAdded = {RN: 0, PCA: 0};

  console.log('Looking for questions for date:', today);

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'M/d/yyyy');

    console.log(`Checking row ${i + 1} [${row[1]}]:`);
    console.log(`Date: ${rowDate}, Role: ${row[11]}, Type: ${row[10]}`);

    if (rowDate === today) {
      try {
        const questionData = {
          questionID: row[1],
          question: row[2],
          options: [row[3], row[4], row[5], row[6], row[7], row[8]].filter(Boolean),
          answer: row[9],
          type: row[10] ? row[10].trim() : '',
          targetRole: row[11],
          points: parseInt(row[12], 10) || 1,
          imageUrl: row[13] ? row[13].trim() : ''
        };

        console.log(`Processing ${questionData.questionID} for ${questionData.targetRole}`);
        console.log('Question data:', JSON.stringify(questionData));

        addQuestionToForm(form, questionData, sections);
        questionsAdded[questionData.targetRole]++;
        
        console.log(`Successfully added ${questionData.targetRole} question`);
        Utilities.sleep(100);
      } catch (e) {
        console.error(`Error processing question in row ${i + 1}:`, e.message);
      }
    }
  }

  console.log('Questions added summary:', questionsAdded);
  logToSheet('Questions Added', 'INFO', 
    `Added ${questionsAdded.RN} RN and ${questionsAdded.PCA} PCA questions`);
}

function addQuestionToForm(form, questionData, sections) {
  console.log(`Adding question ${questionData.questionID} to ${questionData.targetRole} section`);
  
  const targetSection = questionData.targetRole === 'RN' ? sections.rn : sections.pca;
  if (!targetSection) {
    throw new Error(`Target section for ${questionData.targetRole} not found`);
  }

  createAndMoveQuestion(form, questionData, targetSection);
}

function createAndMoveQuestion(form, questionData, section) {
  console.log(`Creating question: ${questionData.questionID}`);
  let item;

  try {
    // Handle image first if present
    if (questionData.imageUrl && questionData.imageUrl.trim() !== "") {
      console.log('Processing image for question');
      try {
        const imageBlob = UrlFetchApp.fetch(questionData.imageUrl).getBlob();
        const imageItem = form.addImageItem()
            .setTitle("")
            .setImage(imageBlob)
            .setAlignment(FormApp.Alignment.CENTER);
        
        Utilities.sleep(500);
        moveItemToSection(form, imageItem, section);
        console.log('Image added successfully');
      } catch (e) {
        console.error('Image processing error:', e.message);
      }
    }

    // Create the question
    switch (questionData.type.toLowerCase()) {
      case 'multiple choice':
        item = form.addMultipleChoiceItem();
        const choices = questionData.options
            .filter(option => option && option.trim() !== "" && option !== "N/A")
            .map(option => item.createChoice(option));
        item.setTitle(questionData.question)
            .setChoices(choices)
            .setRequired(true);
        break;

      case 'multiple select':
      case 'checkbox':
        item = form.addCheckboxItem();
        const checkboxChoices = questionData.options
            .filter(option => option && option.trim() !== "" && option !== "N/A")
            .map(option => item.createChoice(option));
        item.setTitle(questionData.question)
            .setChoices(checkboxChoices)
            .setRequired(true);
        break;

      case 'short answer':
        item = form.addTextItem()
            .setTitle(questionData.question)
            .setRequired(true);
        break;

      default:
        throw new Error(`Unsupported question type: ${questionData.type}`);
    }

    if (form.isQuiz() && item) {
      item.setPoints(questionData.points || 1);
    }

    moveItemToSection(form, item, section);
    console.log(`Question ${questionData.questionID} created and moved to section`);
    
    return item;
  } catch (e) {
    console.error(`Error creating question ${questionData.questionID}:`, e.message);
    throw e;
  }
}

function moveItemToSection(form, item, section) {
  if (!item) {
    console.warn('No item to move');
    return;
  }

  try {
    console.log(`Moving item to section: ${section.getTitle()}`);
    const items = form.getItems();
    const sectionIndex = section.getIndex();
    let targetIndex = sectionIndex + 1;

    // Find the last item in this section
    for (let i = sectionIndex + 1; i < items.length; i++) {
      if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) break;
      targetIndex = i + 1;
    }

    if (targetIndex < items.length) {
      form.moveItem(item.getIndex(), targetIndex);
      console.log('Item moved successfully');
    }
  } catch (e) {
    console.error('Move error:', e.message);
  }
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  ScriptApp.newTrigger('updateDailyQuestions')
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .create();
      
  logToSheet('Trigger Setup', 'SUCCESS', 'Daily trigger created');
}


























































