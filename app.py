from time import sleep
import webbrowser

import pyautogui
import openpyxl
import datetime

import pyperclip



# como posso saber qual dia está mais perto?
def retornaDadosProximaReuniao(dataAtual, workbook):  
    """Retorna os dados da próxima reunião que está a 3 dias ou menos da data atual.""" 
    for sheet in workbook:
        for row in sheet.iter_rows():
            if row[1].value is not None:
                dataPlanilha = row[1].value.strftime('%d/%m/%Y')
                dataPlanilha = datetime.datetime.strptime(dataPlanilha, '%d/%m/%Y').date()
                if (dataPlanilha - dataAtual).days <= 3 and (dataPlanilha - dataAtual).days >= 0:
                    return row
        
        
# print(retornaDadosProximaReuniao(dataAtual, workbook)[1].value)
       
       
def  mensagemDesignacoes(dataAtual, workbook):
    """Retorna a mensagem com as designações para a reunião de hoje."""    
    dadosReuniao = retornaDadosProximaReuniao(dataAtual, workbook)
    return (f'''Para a reunião do dia {dadosReuniao[1].value.strftime("%d\\%m")} temos as designações a seguir:
- Indicador Assistência: *{dadosReuniao[2].value}*
Indicador Porta: *{dadosReuniao[3].value}* 
Microfones Volantes: *{dadosReuniao[4].value.split('/')[0]} e {dadosReuniao[4].value.split('/')[1]}*
Palco: *{dadosReuniao[5].value}*
Som: *{dadosReuniao[6].value}*
Leitor: *{dadosReuniao[7].value}*

        
Para o irmão que não puder comparecer, pedimos que entre em contato com o irmão Junior para que ele possa providenciar um substituto.

Agradecemos a colaboração de todos e que sempre possamos agir de acordo com o que está escrito em 1 Coríntios 14:40 - "Mas que todas as coisas sejam feitas com decência e com ordem".
        ''')
  
# ler planilha
workbook = openpyxl.load_workbook('designacoes.xlsx')

# pegar data atual 
dataAtual = datetime.datetime.now().date()
        
# link para o grupo, Deixando Claro que Tem que estar logado no whatsapp web
url = 'https://web.whatsapp.com'
webbrowser.open(url)
sleep(20)

# Clica para pesquisar no whatsapp
pyautogui.click(250, 240) # Tela 1080x1920 no Opera

# Escrever "Designações" na barra de pesquisa
pyautogui.write('Eu Mesmo')

pyautogui.press('enter')
sleep(3)

message = mensagemDesignacoes(dataAtual, workbook)
lines = message.split('\n')
for line in lines:
    for char in line:
        if char in 'áéíóúâêôãõçàÁÉÍÓÚÂÊÔÃÕÇÀ':
            pyperclip.copy(char)
            pyautogui.hotkey("ctrl", "v")
        else:
            pyautogui.typewrite(char)
    pyautogui.keyDown('shift')
    pyautogui.press('enter')
    pyautogui.keyUp('shift')

seta = pyautogui.locateCenterOnScreen('send.png')
sleep(3)
pyautogui.click(seta[0], seta[1])
sleep(3)
pyautogui.hotkey('ctrl', 'w')

