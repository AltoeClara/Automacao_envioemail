import pandas as pd
import win32com.client as win32
import openpyxl

input_nome = "Aguilar Talita"



tabela = pd.read_excel("ee.xlsx")
format = tabela.iloc[1:10000].query(f'Solicitante== "{input_nome}"')
html_format = format.to_html(index=False)


tabela_email = pd.read_excel("fv.xlsx")
format_email = tabela_email.iloc[0:10].query(f'Solicitante== "{input_nome}"')
input_email = format_email.iloc[0,1]


outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = input_email
email.Subject = "Avaliação Tickets TI"
email.HTMLBody = f"""

Olá, tudo bem?<br>
Identificamos que alguns chamados ainda estão aguardando sua avaliação de atendimento.<br>
Estamos compartilhando a lista com um fácil link de acesso para que você consiga acessá-lo e 
fazer sua avaliação.<br>
Lembramos que sua avaliação é muito importante pois, contribui com nosso processo de melhora 
contínua de nossos atendimentos.<br>
Segue a relação e caso tenha alguma dúvida por favor nos procure.<br>
<br>

{html_format}

<br>

Atenciosamente,<br><br>

<font style="font-size: 30px;"><b>Maria Clara Lopes</b></font><br>
<b>Jovem Aprendiz de TI</b><br>
+55 (12) 3644-8449<br>
maria.silva@adium.com.br
"""
email.Send()
