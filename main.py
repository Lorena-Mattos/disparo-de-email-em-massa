from sendgrid.helpers.mail import Mail
import csv
import sendgrid
import config

apiKey = config.chave_api

# Leitura dos contatos a partir de um arquivo CSV
contatos = []
contatos_detalhes = {}  # Dicionário para armazenar as informações detalhadas de cada contato

with open('contatos.csv', 'r') as file:
    reader = csv.reader(file, delimiter=';')  # Delimitador é o ponto e vírgula (";")
    next(reader)  # Pula o cabeçalho
    for row in reader:
        contatos.append(row)

# Configuração e envio do e-mail para cada contato
for contato in contatos:
    nome = str(contato[0])
    numero = str(contato[1])
    processo = str(contato[2])
    valor = str(contato[3])
    email = str(contato[4])

    if email in contatos_detalhes:
        # Adiciona as informações detalhadas para o contato existente
        contatos_detalhes[email]['detalhes'].append((nome, numero, processo, valor))
    else:
        # Cria um novo item no dicionário para o contato com as informações iniciais
        contatos_detalhes[email] = {'nome': nome, 'detalhes': [(nome, numero, processo, valor)]}

# Configuração e envio do e-mail para cada contato com todas as informações detalhadas
for email, contato in contatos_detalhes.items():
    nome = contato['nome']
    detalhes = contato['detalhes']

    # Construção do HTML com a tabela estilizada
    html_content = f"""
                   <html>
                   <head>
                   <style>
                   table {{
                     border: 1px solid #1C6EA4;
                     background-color: #EEEEEE;
                     width: 100%;
                     text-align: center;
                     border-collapse: collapse;
                   }}
                   table td, table th {{
                     border: 1px solid #AAAAAA;
                     padding: 8px;
                   }}
                   table tbody td {{
                     font-size: 14px;
                     font-weight: bold;
                   }}
                   table thead {{
                     background: #0A3A6A;
                     border-bottom: 2px solid #444444;
                   }}
                   table thead th {{
                     font-size: 16px;
                     font-weight: bold;
                     color: #FFFFFF;
                     text-align: center;
                     border-left: 2px solid #D0E4F5;
                   }}
                   table thead th:first-child {{
                     border-left: none;
                   }}
                   table tfoot td {{
                     font-size: 14px;
                   }}
                   table tfoot .links {{
                     text-align: right;
                   }}
                   table tfoot .links a {{
                     display: inline-block;
                     background: #1C6EA4;
                     color: #FFFFFF;
                     padding: 2px 8px;
                     border-radius: 5px;
                   }}
                   </style>
                   </head>
                   <body>
                   <p>Olá {nome},</p>
                   <p>Aqui estão os detalhes da cobrança:</p>
                   <table>
                     <thead>
                       <tr>
                         <th>Nome</th>
                         <th>Número</th>
                         <th>Processo</th>
                         <th>Valor (R$)</th>
                       </tr>
                     </thead>
                     <tbody>
                       {''.join([f"<tr><td>{detalhe[0]}</td><td>{detalhe[1]}</td><td>{detalhe[2]}</td>"
                                 f"<td>{detalhe[3]}</td></tr>" for detalhe in detalhes])}
                     </tbody>
                   </table>
                   <p>Obrigado!</p>
                   <p>Atenciosamente,<br>Seu Nome</p>
                   </body>
                   </html>
               """
    message = Mail(
        from_email=config.from_email,
        to_emails=email,
        subject='Sending with Twilio SendGrid is Fun',
        html_content=html_content
    )
    # Configuração e envio do e-mail
    try:
        sg = sendgrid.SendGridAPIClient(api_key=apiKey)
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
        print(f'E-mail enviado para {email}')

    except Exception as e:
        print(f'Erro ao enviar e-mail para {email}: {str(e)}')
