import pandas as pd
import smtplib
import email.message

# IMPORTAR BASE DE DADOS:
# a) instalar o openpyxl: pip intall openpyxl no terminal
tabela_vendas = pd.read_excel('Projeto/Vendas.xlsx') # ler e armazenar na variavel o arquivo excel

# VISUALIZAR A BASE DE DADOS:
pd.set_option('display.max_columns', None) # .set_option(opcao, valor) / display.max_columns -> não oculta nenhuma coluna.

print(tabela_vendas)
print('-'*50)

# FATURAMENTO POR LOJA:
# 1º Método: tabela_vendas[['ID Loja', 'Valor Final']] -> filtrando essas duas colunas. Tem que ser entre dois colchetes
# 2º Método: tabela_vendas.groupby('ID Loja').sum() -> agrupando a coluna das lojas e somando os valores das colunas seguinte

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-'*50)

# QUANTIDADE DE PRODUTOS VENDIDOS POR LOJA: 
qtd_prod = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_prod)
print('-'*50)

# TICKET MÉDIO POR PRODUTO EM CADA LOJA: quanto custou em media um valor de um produto vendido
ticket_medio = (faturamento['Valor Final'] / qtd_prod['Quantidade']).to_frame() # .to_frame() -> transformar essa operação em uma tabela.
ticket_medio = ticket_medio.rename(columns = {0 : 'Ticket Médio'}) # renomeia o nome de colunas da tabela.
print(ticket_medio)
print('-'*50)

# ENVIAR UM EMAIL COM O RELATÓRIO:
# importar bibliotecas smtplib e email (gmail)
corpo_email = f"""
    <p>Prezados,</p>
    <p>Segue o relatório de vendas por cada loja.</p>
    <P>Faturamento:</p>
    {faturamento.to_html(formatters = {'Valor Final' : "R${:,.2f}".format})}
    <p>Quantidade vendida:</p>
    {qtd_prod.to_html()}
    <p>Ticket Médio dos produtos em cada loja:</p>
    {ticket_medio.to_html(formatters = {'Ticket Médio' : "R${:,.2f}".format})}
    <p>À disposição.</p>
    """

msg = email.message.Message()
msg['Subject'] = 'Assunto'
msg['From'] = 'remetente'
msg['To'] = 'destinatário'
password = 'senha' # senha do email no "from".
msg.add_header('Content-Type', 'text/html')
msg.set_payload(corpo_email)

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls() # deixa seu código mais seguro para envio
# Login Credentials for sending the mail
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

print('Email enviado')
