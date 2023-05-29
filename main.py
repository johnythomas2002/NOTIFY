import PySimpleGUI as sg
import pandas as pd
from pathlib import Path
from data_Entry import data_win
from check_win import data_check

sg.theme('DarkTeal9')
current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data.xlsx'
df = pd.read_excel(EXCEL_FILE)


def start_win():
    layout1 = [
        [sg.Text('Date App')],
        [sg.Text('Enter Dates', size=(15, 1)), sg.Button('Click', key='Enter', enable_events=True)],
        [sg.Text('Check Todays Dates', size=(15, 1)), sg.Button('Click', key='Check', enable_events=True)],
        [sg.Exit()]
    ]
    return sg.Window('DateAPP', layout1, finalize=True)


window1 = start_win()
window2 = None
window3 = None

while True:
    window, event, values = sg.read_all_windows()
    if event == sg.WIN_CLOSED or event == 'Exit':
        window.close()
        if window == window2:
            window2 = None
        elif window == window3:
            window3 = None
        elif window == window1:
            break
    elif event == 'Enter' and not window2:
        window2 = data_win()
    elif event == 'Check' and not window3:
        window3 = data_check()

window.close()
