# Auto Logging Application Emails into Sheets

## Overview

This is a simple python script that uses Google APIs to read unread emails from your inbox, decode the content, and then pass them into a LLM model to process the information and return the necessary information, and then log it onto Google Sheets.
This script access LLMs through OpenRouter.
## How To Run 

1. Clone or Fork this repo and cd into the repository
2. Create and activate an virtual environment (for isolating your project dependencies) using:
  ```
  In WSL/Linux/macOS:
  python3 -m venv venv
  source venv/bin/activate

  In Windows:
  python -m venv venv
  venv\Scripts\activate
  ```
4. Download the dependencies from requirements.txt using:
   ```
   pip install -r requirements.txt
   ```
5. Create an .env file, this is where you will put your keys such as OpenRouter API Key, Google Sheets ID, etc. It should look something like the following:
    ```
    .env
    OPENAI_API_KEY="YOUR OPEN AI KEY"
    SPREADSHEET_ID="YOUR SPREADSHEET ID"
    ```
6. Run the script using:
   ```
   python main.py
   ```

You can create your OPEN_AI_KEY on OpenRouter and can use a free model: https://openrouter.ai/models?q=free
The current model used in the script is: meta-llama/llama-3.3-70b-instruct:free
If you want to change the model, select any of the models you wish to use and change it on line 211 in main.py.

You can find your Google Sheets ID using: 
```
The Spreadsheet ID is a long string of characters (letters, numbers, hyphens, and underscores) found within the URL. It is located after /d/ and before the next /.
-- Google Gemini
```

To use this script correctly, initiated your Google Sheets looking like the following:

Company	| Position	| Status	| Application Date	| Response Date																					

Click on Column C and select Format and then Conditional Formatting, you can now give your status different colors according to the status. 

Have fun!
   
