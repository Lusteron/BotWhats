import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
webbrowser.open('https://web.whatsapp.com')
sleep(30)
workbook = openpyxl.load_workbook('planilha.xlsx')
paginas_clientes = workbook['Planilha1']

for linha in paginas_clientes.iter_rows(min_row=2):
    #nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value 
    valor = linha[3].value
    mensagem = f'Olá {nome}, sua conta no valor de R$:{valor} vence no dia {vencimento.strftime('%d/%m/%Y')}. favor entrar em contato!'
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        #https://web.whatsapp.com/send?phone=&text¢
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='',encoding='utf-8')as arquivo:
            arquivo.write(f'{nome}, {telefone}')
