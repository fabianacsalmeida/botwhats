import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30)  # Espera o WhatsApp carregar

# Ler a planilha e guardar informações sobre nome e telefone
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

# Loop pelos contatos na planilha
for linha in pagina_clientes.iter_rows(min_row=2):
    # Extrair nome e telefone
    nome = linha[0].value
    telefone = str(linha[1].value)  # Certifique-se de converter o telefone para string

    # Formatar a mensagem
    mensagem = f'Olá {nome}, meu nome é Fabiana. Gostaria de apresentar nossa nova oferta de seguros.'

    # Criar links personalizados do WhatsApp e enviar mensagens
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        
        # Esperar o WhatsApp carregar a página de envio da mensagem
        sleep(10)
        
        # Simular o envio da mensagem pressionando "Enter"
        pyautogui.press('enter')
        
        # Esperar para fechar a aba após o envio
        sleep(15)
        pyautogui.press('esc')  # Pressiona 'ESC' para sair ou pode usar 'ctrl+w' para fechar a aba
        
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}. Erro: {e}')
        
        # Registrar erro no arquivo CSV
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')