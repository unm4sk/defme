# defme - Find definitions easily
## What is it?
Defme is a script written in Python, which allows a user to get definitions for words in Google Spreadsheets. It's pretty easy to install and very easy to use.

## Installation:
1) Install these libraries:
```
pip install rich gspread PyDictionary
```
2) Add a new project on Google Cloud, connecting to **Google Drive API** and **Google Spreadsheets API**. For a more descriptive tutorial on that, please watch this video: `link to be added...`
3) Open `settings.py` and paste:
    * your `.json` file location in `sa_name` variable
    * name of your spreadsheet in `spreadsheet` variable
    * name of worksheet in `worksheet` variable
4) You are ready to go!

## Preview
![Error](resources/defme.gif)

## Notes
Sometimes PyDictionary API returns a `List Out of Range error`. Simple solution: open `core.py` in PyDictionary lib's folder and change `disable_errors = True`. 