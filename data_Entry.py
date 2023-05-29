import PySimpleGUI as sg
import pandas as pd
from pathlib import Path

sg.theme('DarkTeal9')
current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data.xlsx'
df = pd.read_excel(EXCEL_FILE)


def data_win():
    global df
    layout2 = [
        [sg.Text('Please fill out the following fields:')],
        [sg.Text('Name', size=(15, 1)), sg.InputText(key='Name')],
        [sg.Text('Fathers Name', size=(15, 1)), sg.InputText(key='Fathers Name')],
        [sg.Text('Mothers Name', size=(15, 1)), sg.InputText(key='Mothers Name')],
        [sg.Text('House Name', size=(15, 1)), sg.InputText(key='House Name')],
        [sg.Text('Date Of Birth', size=(15, 1)), sg.InputText(key='Date Of Birth')],
        [sg.Text('Wedding Date', size=(15, 1)), sg.InputText(key='Wedding Date')],
        [sg.Submit(), sg.Button('Clear'), sg.Exit()]
    ]
    window = sg.Window('Simple data entry form', layout2,finalize=True)

    def clear_input():
        for key in values:
            window[key]('')
        return None


    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Clear':
            clear_input()
        if event == 'Submit':
            new_record = pd.DataFrame(values, index=[0])
            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Data saved!')
            clear_input()
        if event == 'Date Of Birth':
            name = values['Date Of Birth']
            if name == '':
                continue
            index = df[df.Date_Of_Birth == name].first_valid_index()
            if index is not None:
                lst =  df.iloc[index].to_list()
                for key, value in zip(['Name', 'Fathers Name', 'Mothers Name', 'House Name', 'Date Of Birth', 'Wedding Date'], lst):
                    window[key].update(value)
    window.close()
