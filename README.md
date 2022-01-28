# defme - Find definitions easily
## What is it?
Defme is a script written in Python, which allows a user to get definitions for words in Google Spreadsheets/Microsoft Excel. It's pretty easy to install and very easy to use.

## Installation:
### Defme - Google Spreadsheets version
1) Install these libraries:
    ```
    pip install rich gspread PyDictionary
    ```
2) Add a new project on Google Cloud, connecting to **Google Drive API** and **Google Spreadsheets API**. For a more descriptive tutorial on that, please watch this video: `link to be added...`
3) Open `settings.py` and paste:
    * your `.json` file location in `sa_name` variable
    * name of your spreadsheet in `spreadsheet` variable
    * name of worksheet in `worksheet` variable
4) Execute the script:
    ```
    python defme.py
    ```
5) You are ready to go!
### Defme - Microsoft Excel version
1) Install these libraries:
```
pip install rich gspread PyDictionary
```
2) Download and paste a file location into `settings.py` file    (*ignore* `sa_name` variable). Specify a worksheet you want to work with.    
    Should look like this:
    ```
    spreadsheet = '/path-to-file/file'
    worksheet = 'name-of-worksheet'
    ```
    You **should't** need to add file extension at the end!
3) Run the script:
    ```
    python defme_excel.py
    ```
4) You are ready to go!

## Preview
![Error](resources/defme.gif)

## Notes
While running the script you may see `Error: A Term must be only a single word` messages. Don't worry about them! *It's just the way PyDictionary module works.*