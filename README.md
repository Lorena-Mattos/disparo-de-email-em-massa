<h1 align="center"> Envio de E-mail em Massa 👩‍💻 </h1>

<p align="center">
Projeto feito por Lorena Mattos. <br/>
</p>

<p align="center">
   <a href="#-tecnologias">Tecnologias</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
  <a href="#-projeto">Projeto</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</p>


## 🚀 Tecnologias

Esse projeto foi desenvolvido com as seguintes tecnologias:

- Python
- SendGrid
- openpyxl

## 💻 Projeto

O projeto em questão é responsável por enviar e-mails personalizados para uma lista de contatos, utilizando a biblioteca SendGrid em Python. Ele lê as informações dos contatos de um arquivo CSV, configura e envia o e-mail para cada contato com detalhes específicos da cobrança.

Aqui está uma descrição do que o código faz passo a passo:

Lê os contatos a partir de um arquivo CSV chamado contatos.csv, que deve estar no formato apropriado (cada linha contendo nome, número, processo, valor e e-mail).
Armazena as informações detalhadas de cada contato em um dicionário chamado contatos_detalhes, agrupando os contatos por e-mail.
Configura e envia o e-mail para cada contato usando a biblioteca SendGrid.

Para cada e-mail enviado, o código verifica se o envio foi bem-sucedido. Se for, registra o e-mail na planilha "Sucesso" do arquivo Excel de relatório. Caso contrário, registra o e-mail e a descrição do erro na planilha "Falha".
No final, o código salva o relatório em um arquivo Excel chamado relatorio_emails.xlsx, contendo as informações dos e-mails enviados com sucesso e os que falharam.

O código também inclui a definição das tecnologias usadas, como Python, SendGrid e openpyxl, bem como instruções sobre como configurar e executar o código, incluindo a necessidade de instalar as dependências adequadas e fornecer as informações corretas nos arquivos config.py e contatos.csv.

Em suma, o objetivo desse código é automatizar o envio de e-mails personalizados para uma lista de contatos, fornecendo informações detalhadas de cobrança, e gerar um relatório em Excel que registra quais e-mails foram enviados com sucesso e quais falharam.

## Instruções de Uso 📃

1. Faça o clone deste repositório.
2. Instale as dependências necessárias executando o seguinte comando: `pip install -r requirements.txt`

3. Insira os detalhes dos contatos no arquivo `contatos.csv` no formato apropriado.
4. Configure as informações de e-mail remetente em `config.py`.
5. Execute o script `main.py` para enviar os e-mails.


## Requisitos

Antes de executar este projeto, certifique-se de ter as seguintes dependências instaladas:

- Python 3.x
- SendGrid API Key (obtenha uma em [https://sendgrid.com](https://sendgrid.com))

## Contribuições

Contribuições são sempre bem-vindas! Sinta-se à vontade para abrir issues e pull requests.

## Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE).

 
<p align="center">
Feito com ♥ by Lorena Mattos :wave:
<a href="https://lorena-mattos.github.io/links-da-lorena/">Faça parte das minhas redes!</a>
</p> 

