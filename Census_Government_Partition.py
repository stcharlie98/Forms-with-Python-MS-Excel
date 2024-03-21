from pathlib import Path
import PySimpleGUI as sg
import pandas as pd

sg.theme('LightBlue1')

Excel_Dados = 'Excel_Dados.xlsx'
df = pd.read_excel(Excel_Dados)

layout = [
    [sg.Text('Parte 1 de 6')],
    [sg.Text('Preencha os campos a seguir:')],
    [sg.Text('Permissionária:', size=(15,1)), sg.InputText(key='Permissionária')],
    [sg.Text('Cidade de plantio:', size=(15,1)), sg.InputText(key='Cidade')],
    [sg.Text('Estado:', size=(15,1)),  
                            sg.Radio('SP', group_id='estado', key='SP', enable_events=True),
                            sg.Radio('MG', group_id='estado', key='MG', enable_events=True),
                            sg.Radio('PR', group_id='estado', key='PR', enable_events=True),
                            sg.Radio('Outro', group_id='estado', key='Outro_to_Hide', enable_events=True),
                            sg.InputText(key='Outro', size=(5,1), enable_events=True)],
    [sg.Text('Permissionária de:', size=(15,1)),
                            sg.Checkbox('Box(es)', key='Galpão Permanente (GP)'),
                            sg.Checkbox('Módulo(s)', key='Mercado Livre (ML)'),
                            sg.Checkbox('Loja(s)', key='Loja (LJ)')],
    [sg.Text('GP(s):', size=(15,1)), 
                            sg.Checkbox('01', key='GP01'),
                            sg.Checkbox('02', key='GP02'),
                            sg.Checkbox('03', key='GP03'),
                            sg.Checkbox('04', key='GP04'),
                            sg.Checkbox('PABC', key='ABC')],
    [sg.Text('ML(s):', size=(15,1)), 
                            sg.Checkbox('01', key='M1'),
                            sg.Checkbox('Central', key='Central'),
                            sg.Checkbox('02', key='M2'),
                            sg.Checkbox('04', key='M4')],
    [sg.Text('Loja(s):', size=(15,1)), sg.Spin([i for i in range(1,10)],
                                            initial_value='', key='LJ', size=(3, 1))],
    [sg.Submit('Enviar'), sg.Button('Limpar'), sg.Exit('Sair')]
]

window = sg.Window('Censo Ceasa 2024', layout, resizable=True, disable_close=True)

def clear_input():
    for key in values:
        window[key]('')
    return None

exit_confirmed = False 

while True:
    event, values = window.read()
   
    if event == 'Enviar':
        df = pd.concat([df, pd.DataFrame(values, index=[0])], ignore_index=True)
        df.to_excel(Excel_Dados, index=False) 
        sg.popup('Salvo com sucesso!')
        clear_input()

    if event == 'Limpar':
        clear_input()

    if event == 'Sair':
        if sg.popup_yes_no('Você tem certeza de que quer sair? Todos os dados serão perdidos.') == 'Sim': 
            break
                    
window.close()