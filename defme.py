from rich.console import Console
from rich.progress import track
from rich.theme import Theme
import gspread
import os
import time
from settings import spreadsheet, worksheet, sa_name
from PyDictionary import PyDictionary

dictionary = PyDictionary()

custom_theme = Theme({'success': 'bold underline green', 'warning': 'underline yellow', 'error': 'bold red', 'error_api': 'blue'})
console = Console(theme=custom_theme)

os.system('cls' if os.name == 'nt' else 'clear')

#TODO implement undrelines if word appears in definition
#TODO add synonyms & antonyms
#TODO add examples <- create my own method 

service_account = gspread.service_account(filename=sa_name)
spread_sheet = service_account.open(spreadsheet)
work_sheet = spread_sheet.worksheet(worksheet)

possibles = ['Noun', 'Verb', 'Adjective', 'Adverb']
not_single = []
review = []
cur_pos = 2
rows_count = 1
ran = False

def county():
    global rows_count, ran
    cnt_track = 0
    for i in range(rows_count, work_sheet.row_count+1):
        if cnt_track >=10: # checking for the end of spreadsheet
            rows_count -= 10
            break
        if (work_sheet.cell(i, 1).value == None):
            cnt_track += 1
            rows_count += 1
        elif (work_sheet.cell(i, 1).value != None):
            cnt_track = 0
            rows_count += 1
        time.sleep(0.1)
    console.print('Words counted!', style='success')
    ran = True
    return rows_count-1


def dict(i):
    result = ''
    definition = dictionary.meaning(f'{work_sheet.cell(i, 1).value}', disable_errors=True)
    # print(definition)
    for i in possibles:
        try: 
            definition[i]
            if len(definition[i]) == 1:
                result += f'{i}: ' + definition[i][0] + '\n'
            elif len(definition[i]) >=3:
                result += f'{i}: '
                for pos in range(3):
                    result += f'{pos+1}) ' + definition[i][pos] + '\n'
            else:
                result += f'{i}: '
                for pos in range(len(definition[i])+1):
                    result += f'{pos+1}) ' + definition[i][pos] + '\n'
        except:
            pass

    return result.strip()

def search(beg=2,end=rows_count): # current position
    global cur_pos, review
    for i in track(range(beg, rows_count), description='Processing...'):
        try: 
            if len((work_sheet.cell(i, 1).value).split()) > 1:
                not_single.append(f'A{i}')
        except:
            pass
        if (work_sheet.cell(i, 1).value != None) and (work_sheet.cell(i, 2).value == None):
            work_sheet.update(f'B{i}', dict(i))
        elif (work_sheet.cell(i, 1).value == None) and (work_sheet.cell(i, 2).value != None): 
            work_sheet.update(f'A{i}', 'REVIEW!!!')
            review.append(f'A{i}:B{i}')
            # print it using rich 
        cur_pos += 1
        time.sleep(0.1)

def main():
    global ran
    if not ran:
        with console.status('Counting all words in cells...', spinner='material'):
            county()
        with console.status('Cooldown', spinner='clock'):
            time.sleep(5)
    search(beg=cur_pos)
    console.print('Program sucessfully executed!', style='success')
    if review:
        console.print('\nDuring the execution these cells missing word but having definition were found:\n', style='warning')
        print(', '.join(review))
    if not_single:
        console.print('\nDuring the execution these cells containing multiple words were found:\n', style='warning')
        print(', '.join(not_single))


if __name__ == '__main__':
    while True:
        try:
            main()
            break
        except gspread.exceptions.APIError as err:
            console.print('Error on the API side occured!', style='error_api')
            with console.status("Google API's has some restrictions! We're sorry!", spinner='monkey'):
                time.sleep(20)
            # print(err)
        except gspread.exceptions.CellNotFound as err:
            console.print('Cell was not found!', style='error')
            pass
        except gspread.exceptions.WorksheetNotFound as err:
            console.print('Worksheet does not exist or not acessible!', style='error')