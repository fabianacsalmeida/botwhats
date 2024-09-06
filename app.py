import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

# 1. Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(35)  # Espera mais tempo para garantir o carregamento da página

# 2. Ler a planilha e obter informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

# 3. Loop sobre os contatos da planilha
for linha in pagina_clientes.iter_rows(min_row=2):
    # Extrair nome, telefone e data de vencimento
    nome = linha[0].value
    telefone = str(linha[1].value)  # Certifique-se de que o telefone é uma string
    vencimento = linha[2].value
    
    # Formatar a mensagem personalizada
    mensagem = (
        f'Olá {nome}, 💰 Você quer poupar aquela graninha extra? Invista de 100 a 200 reais no Ourocap! 💰'
        '**Guarde seu dinheiro e concorra a 3.380 prêmios**! 😱 São 70 prêmios de até R$ 20 mil todo mês + sorteios semestrais e o especial anual de R$ 2,5 milhões. 🎉'
        'Próximo sorteio: 11/09.'
        '#RealizeComOurocap'
        '#SeuEuDoFuturoAgradece 🌟 📲')
        
    # Criar o link de mensagem personalizada do WhatsApp
    try:
        # Verifique se o telefone tem o formato correto com código do país
        if not telefone.startswith("+"):
            telefone = "+55" + telefone  # Adiciona o código do Brasil por padrão

        # Criar o link de envio
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)

        # Aguardar para o WhatsApp carregar a conversa
        sleep(20)

    

        # Fechar a aba atual do navegador
        pyautogui.hotkey('enter')
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
  
        
    except Exception as e:
        print(f'Erro ao enviar mensagem para {nome}: {e}')

        # Salvar erros no arquivo CSV
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
