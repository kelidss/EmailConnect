from imap_tools import AND, MailBox  # type: ignore
from openpyxl import Workbook

usuario = 'EMAIL'
senha = 'KEY'
meu_email = MailBox("imap.gmail.com").login(usuario, senha)

listas_email = meu_email.fetch(AND(from_="=EMAIL"))

wb = Workbook()
ws = wb.active

for email in listas_email:
    print(f"De: {email.from_}, Assunto: {email.subject}, Data: {email.date}")
    print(email.text)
    ws.append([email.from_, email.subject, email.date.strftime(  # type: ignore
        "%Y-%m-%d %H:%M:%S"), email.text])
    # possivel erro (use caso queira salvar tudo na mesma célula)
    # ws['A1'] = email.text
wb.save("texto_email.xlsx")
meu_email.logout()
