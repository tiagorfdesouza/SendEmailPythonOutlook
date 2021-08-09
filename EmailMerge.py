import pandas as pd
import win32com.client as win32
import time

file = pd.ExcelFile('Excel Workbook Template.xlsx')

df = file.parse('Sheet1')

for index, row in df.iterrows():
    email = (row['Email Address'])
    subject = (row['Subject'])
    #body = str((row['Email HTML Body']))

    if pd.isnull(email) or pd.isnull(subject):
        continue

    olMailItem = 0x0
    obj = win32.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.HTMLBody = """Prezado Cliente,<br>

<p>Solicitamos o envio das documentações, arquivos e informações fiscais até o dia 5 (Cinco), 
de modo a atendermos os prazos acordados pelos devidos órgãos, evitando atrasos nas entregas dos impostos, 
declarações acessórias, pagamentos de multas, descredenciamentos de incentivos entre outros.</p>

<p>O fechamento é realizado por ordem do envio da documentação, lembrando que documentação 
enviada próximo do vencimento da obrigação é de inteira responsabilidade do contribuinte, 
pois não teremos tempo hábil para uma análise analítica.</p><br>

<p>OBS:</p> 

<p>Por gentileza, enviar mensalmente para o Departamento Fiscal os relatórios consolidados 
dos extratos referente as aplicações financeiras de todos os bancos, caso não haja nenhuma aplicação 
financeira no recorrente mês, informar por e-mail.</p><br><br>

<p>Segue em anexo Checklist da documentação solicitada.</p><br>

Atenciosamente, Departamento Fiscal!"""
    newMail.To = email
    anexo = "C:\imagens\Checklist - Regime Normal.pdf"
    newMail.Attachments.Add(anexo)
    time.sleep(5)
    newMail.Send()
    print("Email enviado Com Sucesso!!!")