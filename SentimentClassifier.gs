/**
 * Sentiment Classifier using OpenAI API
 * 
 * This script allows you to classify user feedback in Google Sheets as positive, neutral, or negative.
 * It uses the OpenAI API with a customizable system prompt stored in a cell.
 */

// API Configuration
const OPENAI_API_KEY_CELL = "B1"; // Cell containing your OpenAI API key
const SYSTEM_PROMPT_CELL = "B2"; // Cell containing your customizable system prompt
const MODEL_CELL = "B3"; // Cell containing the model name (e.g., "gpt-3.5-turbo")

// Create a menu when the spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sentiment Analysis')
    .addItem('Classify Selected Feedback', 'classifySelectedFeedback')
    .addItem('Classify All Feedback', 'classifyAllFeedback')
    .addSeparator()
    .addItem('Setup Instructions', 'showSetupInstructions')
    .addToUi();
}

/**
 * Show setup instructions in a modal dialog
 */
function showSetupInstructions() {
  const ui = SpreadsheetApp.getUi();
  const instructions = 
    "Setup Instructions:\n\n" +
    "1. Enter your OpenAI API key in cell " + OPENAI_API_KEY_CELL + "\n" +
    "2. Enter your system prompt in cell " + SYSTEM_PROMPT_CELL + " (example provided)\n" +
    "3. Enter the model name in cell " + MODEL_CELL + " (e.g., 'gpt-3.5-turbo')\n" +
    "4. Format your sheet with feedback text in column A starting from row 5\n" +
    "5. The sentiment classification will appear in column B\n\n" +
    "Example system prompt:\n" +
    "You are a sentiment classifier. Analyze the user feedback and classify it as 'positive', 'neutral', or 'negative'. Only respond with one of these three words.";
  
  ui.alert("Setup Instructions", instructions, ui.ButtonSet.OK);
}

/**
 * Classify the selected feedback cells
 */
function classifySelectedFeedback() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  
  if (selection.getColumn() != 1) {
    SpreadsheetApp.getUi().alert("Please select cells in column A containing feedback to classify.");
    return;
  }
  
  // Get configuration
  const apiKey = sheet.getRange(OPENAI_API_KEY_CELL).getValue();
  const systemPrompt = sheet.getRange(SYSTEM_PROMPT_CELL).getValue();
  const model = sheet.getRange(MODEL_CELL).getValue();
  
  if (!apiKey || !systemPrompt || !model) {
    showSetupInstructions();
    return;
  }
  
  // Process each selected cell
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    const feedback = sheet.getRange(currentRow, 1).getValue();
    
    if (feedback) {
      const sentiment = classifySentiment(feedback, apiKey, systemPrompt, model);
      sheet.getRange(currentRow, 2).setValue(sentiment);
    }
  }
}

/**
 * Classify all feedback in the sheet
 */
function classifyAllFeedback() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get configuration
  const apiKey = sheet.getRange(OPENAI_API_KEY_CELL).getValue();
  const systemPrompt = sheet.getRange(SYSTEM_PROMPT_CELL).getValue();
  const model = sheet.getRange(MODEL_CELL).getValue();
  
  if (!apiKey || !systemPrompt || !model) {
    showSetupInstructions();
    return;
  }
  
  // Get all data range
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Start from row 5 (adjust as needed)
  const startRow = 5;
  
  for (let i = startRow - 1; i < values.length; i++) {
    const feedback = values[i][0];
    
    if (feedback && typeof feedback === 'string') {
      const sentiment = classifySentiment(feedback, apiKey, systemPrompt, model);
      sheet.getRange(i + 1, 2).setValue(sentiment);
    }
  }
}

/**
 * Classify a single piece of feedback using the OpenAI API
 */
function classifySentiment(feedback, apiKey, systemPrompt, model) {
  try {
    const url = 'https://api.openai.com/v1/chat/completions';
    
    const payload = {
      model: model,
      messages: [
        {
          role: "system",
          content: systemPrompt
        },
        {
          role: "user",
          content: feedback
        }
      ],
      temperature: 0.3 // Lower temperature for more consistent results
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.error) {
      Logger.log("API Error: " + responseData.error.message);
      return "ERROR: " + responseData.error.message;
    }
    
    return responseData.choices[0].message.content.trim();
  } catch (error) {
    Logger.log("Error: " + error);
    return "ERROR: " + error.toString();
  }
}

/**
 * Creates an example template in the current sheet
 */
function setupExampleTemplate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear the sheet
  sheet.clear();
  
  // Set up header cells
  sheet.getRange("A1").setValue("OpenAI API Key:");
  sheet.getRange("A2").setValue("System Prompt:");
  sheet.getRange("A3").setValue("Model Name:");
  
  // Set up example system prompt
  sheet.getRange("B2").setValue("You are a sentiment classifier. Analyze the user feedback and classify it as 'positive', 'neutral', or 'negative'. Only respond with one of these three words.");
  sheet.getRange("B3").setValue("gpt-3.5-turbo");
  
  // Set up column headers
  sheet.getRange("A4").setValue("User Feedback");
  sheet.getRange("B4").setValue("Sentiment");
  
  // Format headers
  sheet.getRange("A1:B4").setFontWeight("bold");
  sheet.setColumnWidth(1, 400); // Set column A width
  sheet.setColumnWidth(2, 150); // Set column B width
  
  // Add some example feedback
  const examples = [
    "I absolutely love this product! It has made my life so much easier.",
    "The product works as expected. Nothing special.",
    "This is the worst experience I've ever had. Complete waste of money.",
    "It's okay I guess, but I expected more features for the price.",
    "Not bad, does what it's supposed to do."
  ];
  
  for (let i = 0; i < examples.length; i++) {
    sheet.getRange(i + 5, 1).setValue(examples[i]);
  }
}
