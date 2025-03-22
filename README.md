# Google Sheets Sentiment Classifier

Easily classify user feedback in Google Sheets as positive, neutral, or negative using the OpenAI API with a customizable system prompt.

## Features

- Classify single or multiple feedback entries with one click
- Customize your system prompt directly in a spreadsheet cell
- Choose which LLM model to use
- Simple setup with no external dependencies

## Setup Instructions

### 1. Create a new Google Sheet

Start by creating a new Google Sheet or opening an existing one where you want to analyze feedback.

### 2. Add the Apps Script

1. Click on **Extensions > Apps Script** in your Google Sheets menu
2. Delete any code in the script editor
3. Copy and paste the entire code from [SentimentClassifier.gs](SentimentClassifier.gs) into the editor
4. Save the project (Ctrl+S or ⌘+S) and give it a name like "Sentiment Classifier"
5. Close the Apps Script editor and refresh your Google Sheet

### 3. Configure Your Sheet

1. Set up your sheet with the following information:
   - Cell B1: Your OpenAI API key
   - Cell B2: Your system prompt (see example below)
   - Cell B3: The model name (e.g., "gpt-3.5-turbo")
2. From row 5 onwards, place your user feedback in column A

### 4. Example System Prompt

Here's an example system prompt you can use in cell B2:

```
You are a sentiment classifier. Analyze the user feedback and classify it as 'positive', 'neutral', or 'negative'. Only respond with one of these three words.
```

You can customize this prompt as needed to get different classification schemes or add specific instructions for your use case.

## Usage

After setup, you'll see a new menu item called "Sentiment Analysis" in your Google Sheets menu. It offers these options:

- **Classify Selected Feedback**: Analyzes only the cells you've selected in column A
- **Classify All Feedback**: Analyzes all feedback entries in column A (starting from row 5)
- **Setup Instructions**: Shows a quick reminder of how to configure the sheet

The classification results will appear in column B next to each feedback entry.

## Example Classification Results

| User Feedback | Sentiment |
|---------------|-----------|
| I absolutely love this product! It has made my life so much easier. | positive |
| The product works as expected. Nothing special. | neutral |
| This is the worst experience I've ever had. Complete waste of money. | negative |
| It's okay I guess, but I expected more features for the price. | neutral |

## Customization Options

You can modify your system prompt in cell B2 to get different classification behaviors:

- **Different Categories**: Change from positive/neutral/negative to other categories like "urgent/non-urgent" or "bug/feature-request/question"
- **More Detailed Analysis**: Request more information like sentiment strength or specific aspects mentioned
- **Special Instructions**: Add domain-specific context or rules for classification

## Troubleshooting

If you encounter issues:

1. Make sure your OpenAI API key is valid and correctly entered
2. Check that you've entered a model name that exists (e.g., "gpt-3.5-turbo" or "gpt-4")
3. Ensure your system prompt is clear and properly formatted

## Notes

- API calls consume OpenAI credits according to your account's pricing plan
- For large volumes of feedback, consider batching your classification tasks

## License

MIT License - feel free to modify and use as needed!
