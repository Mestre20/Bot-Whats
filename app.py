import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Ler planilha e guardar informações sobre nome, telefone e nome do técnico
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Clientes']

for linha in pagina_clientes.iter_rows(min_row=2):
    # Nome, Telefone, Técnico
    nome = linha[0].value
    telefone = linha[1].value
    tecnico = linha[2].value

    # Verificar se alguma das colunas está vazia
    if nome is None or telefone is None or tecnico is None:
        print(f'Coluna vazia na linha {linha[0].row}. Pulando para próxima linha...')
        continue
    
    mensagem = f'Bom dia {nome}! O técnico {tecnico} vai instalar hoje na sua residência. Fique de olho no celular.'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(30)
        
        # Enviar a mensagem usando pyautogui
        pyautogui.typewrite('\n')  # Enviar a mensagem
        sleep(2)

        # Fechar a janela do navegador
        pyautogui.hotkey('ctrl','w')
        sleep(2)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')

# Fechar o navegador
webbrowser.close()
