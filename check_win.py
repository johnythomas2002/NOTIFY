import PySimpleGUI as sg
import pandas as pd
from pathlib import Path

sg.theme('DarkTeal9')
current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data.xlsx'
df = pd.read_excel(EXCEL_FILE)

def data_check():
    layout3 = [
        [sg.Text("Enter the date to check for birthdays:")],
        [sg.Input(key='Date'), sg.Button('Check', key='Check')],
        [sg.Multiline(size=(60, 15), key='Output', disabled=True)]
    ]

    window = sg.Window("Check Birthdays", layout3, finalize=True)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break

        if event == 'Check':
            date = values['Date']

            if date:
                birthdays = df[df['Date Of Birth'] == date]

                if not birthdays.empty:
                    output = 'Birthdays on {}:\n\n'.format(date)
                    for _, row in birthdays.iterrows():
                        output += "Name: {}\n".format(row['Name'])
                        output += "Father's Name: {}\n".format(row['Fathers Name'])
                        output += "House Name: {}\n\n".format(row['House Name'])
                else:
                    output = 'No birthdays on {}'.format(date)

                window['Output'].update(output)

    window.close()
