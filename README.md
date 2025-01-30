<h1>Sistema para envio automático de E-mails e relação de usuário em Planilha Excel</h1>

> Status: Concluído


<h2>Bibliotecas</h2>

Para o funcionamento do projeto, seguir a instalação das bibliotecas:

```
pip install pandas
pip install pywin32
pip install openpyxl
```

<h2>Planilha Excel</h2>

Para que o sistema colete os dados de um usuário e os envie para o e-mail correspondente, as colunas da tabela do Excel devem estar formatadas da seguinte maneira:

ID | Solicitante | Assunto | Link
-- | ----------- | ------- | ----
000001 | Nome do Usuário 1 | Título do ticket 1 | https://linkdoticket1.com
000002 | Nome do Usuário 2 | Título do ticket 2 | https://linkdoticket2.com
000003 | Nome do Usuário 3 | Título do ticket 3 | https://linkdoticket3.com
000004 | Nome do Usuário 4 | Título do ticket 4 | https://linkdoticket4.com
