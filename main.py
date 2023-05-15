from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime
layout = [[sg.Text('ID сотрудника'), sg.Push(), sg.Input(key='master')], [sg.Text('читательский билет посетителя'), sg.Push(), sg.Input(key='client')], [sg.Text('ID экземпляра'), sg.Push(), sg.Input(key='book')], [sg.Button('Добавить'), sg.Button('Закрыть')]]
window = sg.Window('База данных библиотеки', layout, element_justification='center')
while True:
 event, values = window.read()
 if event == sg.WIN_CLOSED or event == 'Закрыть':
    break
 if event == 'Добавить':
  try:
   wb = load_workbook('bd.xlsx')
   sheet = wb['Лист1']
   sheet['A1'] = "№"
   sheet['B1'] = "ID сотрудника"
   sheet['C1'] = "читательский билет посетителя"
   sheet['D1'] = "ID экземпляра"
   sheet['E1'] = " Дата и Время добавления в БД"
   ID = len(sheet['ID']) + 1
   time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
   data = [ID, values['master'], values['client'], values['book'], time_stamp]
   sheet.append(data)
   wb.save('bd.xlsx')
   window['master'].update(value='')
   window['client'].update(value='')
   window['book'].update(value='')
   window['master'].set_focus()
   sg.popup('Данные сохранены')

  except PermissionError:
   sg.popup('File in use', 'File is being used by anotherser.\nPlease try again later.')
#fwfwfwfwfwfffwfwfwfwfwf
