# Welcome to iTask 1.2

The iTask project is a project of mine, where the idea was to create something from scratch to understand all the logic of and study the Python programming language and its dependencies.
The premise of the project came from helping my son manage to organize his payments for tasks and his punish (which are punishments applied in relation to what he earns for fulfilling something within our daily life).

As his "allowance" works, he has chores to do at home: help take out the garbage, help take care of the pets, get good grades at school, pass the year, do the dishes, help keep the house clean and help clean the house on cleaning days. Each task has its own value, and today he is putting the information together in an excel spreadsheet, but iTask will help him to have a more interactive view and help organize his financial life (a 15-year-old teenager.)

So thinking about that came the idea for my study, to create something that could help him and bring a challenge for me to be able to create and learn.

To be able to create the .exe file, just use the following command in the terminal in the root folder of the project:
pyinstaller --onefile --windowed --icon='c://path/file.ico' iTask.py

## News...
Now the export to excel function, in addition to saving the spreadsheet locally, also sends a summary and the file by email.

Check requirements.txt file for required libraries.

## YouTube Video (Version 1.0): https://youtu.be/tF78CP-x504
## YouTube Video (Version 1.1): https://youtu.be/b7lem1wA37s
## YouTube Video (Version 1.2): https://youtu.be/uFQaaJ4odgI

# iTask screens

## New screens
![News screen](image/env_exp.png)
![News screen](image/excel_save.png)
![News screen](image/plan_01.png)
![News screen](image/plan_02.png)
![News screen](image/email_send.png)
![News screen](image/email.png)

## Welcome tab

![Welcome screen](image/tab_welcome.png)
![Welcome Cards screen](image/cards.png)

When opening the application, a welcome message is shown and three cards with the summed values of tasks, punish and how much you have to receive at the end of the cycle of tasks (think it would be per month, from the first day to the last day of the month chain).
The cards are updated as the application is opened, to reflect the values it takes from the database the information of the sum of the value column of each database table.

## Tasks Tab

![Task screen](image/tab_task.png)

This screen is where task records take place, there is a simple form to insert the task title (text type field), date (date type field with mask) and value (value type field and accepts decimal numbers).

Button to send form data to the table.

We also have the part where the table is created (still without formatting, but this will be improved over time, in the next versions).

Button to save the table in the database and finally the button to export it to an Excel file.
The export button has a double function, in addition to exporting it cleans the bank, to start a new cycle of tasks and punish.

## Punish tab

![Punish screen](image/tab_punish.png)

This screen is where punish are registered, there is a simple form to insert the title of the fine (text type field), date (date type field with mask) and value (value type field and accepts decimal numbers).

Button to send form data to the table.

We also have the part where the table is created (still without formatting, but this will be improved over time, in the next versions).

Button to save the table in the database and finally the button to export it to an Excel file.
The export button has a double function, in addition to exporting it cleans the bank, to start a new cycle of tasks and punish.

## Report tab

![Report screen](image/report.png)

A tab to show the total amount of tasks and punish in a bar graph.

## Database

![Database screen](image/database.png)
![Database screen](image/database2.png)

Database created in SQLite3 containing two tables, the file is created automatically when opening the application for the first time.

# Bem-vindo ao iTask 1.2

O projeto iTask ?? um projeto de autoria minha, onde a ideia foi criar algo do zero para entender toda a l??gica de e estudar a linguagem de programa????o Python e suas depend??ncias.
A premissa do projeto surgiu em ajudar meu filho a conseguir organizar seus pagamentos por tarefas e suas multas (que s??o castigos aplicados em rela????o ao que ele ganha por cumprir algo dentro do nosso cotidiano).

Como funciona a "mesada" dele, ele tem tarefas de casa a serem cumpridas: ajudar a retirar o lixo, ajudar cuidar dos animais de estima????o, ter notas boas na escola, passar de ano, lavar lou??a, ajudar a manter a casa limpa e ajudar a limpar a casa nos dias de faxina. Cada tarefa tem seu valor, e ele vai unindo hoje a informa????o em uma planilha de excel, por??m o iTask vai ajuda-lo a ter uma vis??o mais interativa e auxiliar nessa fun????o de organizar a vida financeira dele (um adolecente de 15 anos.)

Ent??o pensando nisso que veio a ideia para meu estudo, criar algo que pudesse ajuda-lo e trouxesse um desafio para eu poder criar e aprender.

Para poder criar o arquivo .exe basta usar o seguinte comando no terminal na pasta raiz do projeto:
pyinstaller --onefile --windowed --icon='caminho do arquivo.ico' iTask.py

## Novidades
Agora a fun????o de exportar para o excel al??m de salvar a planilha localmente, tamb??m enviar por e-mail um resumo e o arquivo.

Verificar arquivo requeriments.txt para bibliotecas necess??rias.

# Telas do iTask

## Novas telas
![News screen](image/env_exp.png)
![News screen](image/excel_save.png)
![News screen](image/plan_01.png)
![News screen](image/plan_02.png)
![News screen](image/email_send.png)
![News screen](image/email.png)

## Aba de Boas-vindas

![Welcome screen](image/tab_welcome.png)
![Welcome Cards screen](image/cards.png)

Ao abrir o aplicativo ?? mostrado uma mensagem de boas vindas e tr??s cards com os valores somados de tarefas, multas e o quanto tem a receber ao final do ciclo de tarefas (pense que seria por m??s, do primeiro dia at?? o ultimo dia do m??s corrente).
Os cards s??o atualizados conforme o aplicativo ?? aberto, para refletir os valores ele pega a partir do banco de dados a informa????o da soma da coluna valor de cada tabela do banco de dados.

## Aba de Tarefas

![Task screen](image/tab_task.png)

Nessa tela ?? onde acontecem os registros de tarefas, h?? um formul??rio simples para inserir o titulo da tarefa(campo tipo texto), data (campo tipo data com mascara) e o valor (campo do tipo valor e aceita numeros decimais).

Bot??o para enviar os dados do formul??rio para a tabela.

Temos tamb??m a parte onde ?? criada a tabela (ainda esta sem formata????o, mas isso ser?? melhorado com o tempo, nas pr??ximas vers??es).

Bot??o para salvar a tabela no banco de dados e por ultimo o bot??o para exportar para o arquivo Excel.
O bot??o para exportar tem uma fun????o dupla, al??m de exportar ele limpa o banco, para se iniciar um novo ciclo de tarefas e multas.

## Aba de Multas

![Punish screen](image/tab_punish.png)

Nessa tela ?? onde acontecem os registros de multas, h?? um formul??rio simples para inserir o titulo da multa(campo tipo texto), data (campo tipo data com mascara) e o valor (campo do tipo valor e aceita numeros decimais).

Bot??o para enviar os dados do formul??rio para a tabela.

Temos tamb??m a parte onde ?? criada a tabela (ainda esta sem formata????o, mas isso ser?? melhorado com o tempo, nas pr??ximas vers??es).

Bot??o para salvar a tabela no banco de dados e por ultimo o bot??o para exportar para o arquivo Excel.
O bot??o para exportar tem uma fun????o dupla, al??m de exportar ele limpa o banco, para se iniciar um novo ciclo de tarefas e multas.

## Aba Relat??rio

![Report screen](image/report.png)

Uma aba para mostrar o valor total das tarefas e multas em um gr??fico de barras.

## Banco de dados 

![Database screen](image/database.png)
![Database screen](image/database2.png)

Banco de dados criado em SQLite3 contendo duas tabelas, o arquivo ?? criado automaticamente ao abrir a primeira vez o aplicativo.

## YouTube Video (Version 1.0): https://youtu.be/tF78CP-x504
## YouTube Video (Version 1.1): https://youtu.be/b7lem1wA37s
## YouTube Video (Version 1.2): https://youtu.be/uFQaaJ4odgI

## Quer ajudar o canal? Basta fazer uma doa????o ??????
## Want to help the channel? Just make a donation ??????
# Cryptocurrency 
### Ethereum Mainnet - ETH, USDT >>> Outra criptomoeda para rede / Another cryptocurrency for networking <<<
### Binance Smart Chain - BNB, BUSD >>> Outra criptomoeda para rede / Another cryptocurrency for networking <<<
### Polygon MATIC, wETH, USDC >>> Outra criptomoeda para rede / Another cryptocurrency for networking <<<

## Wallet: 0x8f424A6360BAf9d4111994ad2fDb0adC25004Eb3

# Binance Pay
## Pay ID: 240356717

# PayPal
## dgarcia.saltori@me.com

# PIX >>> Apenas para o Brasil - Only for Brazil <<<
## Chave PIX: dgarcia.saltori@gmail.com 

Meu YouTube: https://youtube.com/@diego.saltori
Meu Instagram: http://instagram.com/diego.saltori
Meu LinkedIn: https://www.linkedin.com/in/dgarciasaltori/