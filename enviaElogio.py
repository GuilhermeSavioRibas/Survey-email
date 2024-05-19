import openpyxl
from html2image import Html2Image
from PIL import Image
import base64
import win32com.client as win32
import os

def ler_dados_excel():
    arquivo_excel = r'C:\Users\Ribas\Desktop\Pasta1.xlsx'
    wb = openpyxl.load_workbook(arquivo_excel)
    planilha = wb.active
    dados = []
    for row in planilha.iter_rows(min_row=2, values_only=True):
        if row[5] == "Não":  # Verifica se a coluna F (índice 5) é "Não"
            dados.append(row)
    return dados

def construir_email(nome_analista, nome_usuario, numero_chamado, mensagem_elogio):
    html_template = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #ffffff;
            color: #FFFFFF;
            overflow: hidden; /* Remove qualquer overflow */
        }}

        .container {{
            background-color: #003134;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            max-width: 600px; /* Define a largura máxima para o container */
            margin: 0 auto; /* Centralize o container horizontalmente */
            position: absolute; /* Posicione absolutamente o container */
        }}

            .header {{
                text-align: center;
                margin-bottom: 20px;
            }}

            .header h1 {{
                font-size: 28px;
                color: #00E28B;
                margin: 0;
                padding: 0;
                text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
            }}  

            .content {{
                font-size: 16px;
                line-height: 1.6;
                color: #FFFFFF;
                padding-top: 10px;
            }}

            .content p {{
                margin: 10px 0;
            }}

            .highlight {{
                font-weight: bold;
            }}

            .highlight-box {{
                width: calc(100% - 20px);
                padding: 10px;
                background-color: rgba(0, 0, 0, 0.2);
                border-radius: 8px;
                margin-top: 20px;
            }}

            .message-box {{
                padding: 20px;
                background-color: rgba(0, 0, 0, 0.2);
                border-radius: 8px;
            }}

            .message-content {{
                font-style: italic;
                color: #FFFFFF;
                margin-bottom: 10px;
            }}

            .stars {{
                font-size: 24px;
                color: #00E28B;
            }}

            .star {{
                margin-right: 5px;
            }}

            .footer {{
                text-align: center;
                margin-top: 20px;
                font-size: 14px;
                color: #FFFFFF;
            }}

            .logos {{
                display: flex;
                justify-content: space-between;
                margin: 20px;
                margin-bottom: -20px;
                align-items: center;
            }}

            .logos img {{
                height: 50px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Reconhecimento de Excelente Atendimento</h1>
            </div>
            <div class="content">
                <p>Olá <span class="highlight">{}</span>,</p>
                <p>Estamos felizes em informar que você recebeu um feedback positivo de um usuário que você atendeu recentemente. Parabéns pelo excelente trabalho!</p>
            </div>
            <div class="highlight-box">
                <div class="message-box">
                    <div class="stars">
                        <span class="star">★</span>
                        <span class="star">★</span>
                        <span class="star">★</span>
                        <span class="star">★</span>
                        <span class="star">★</span>
                    </div>
                    <p><strong>Nome do Usuário:</strong> <span class="highlight">{}</span></p>
                    <p class="message-content">Mensagem do Elogio: <em>"{}"</em></p> 
                    <p><strong>Número do Chamado:</strong> <span class="highlight">{}</span></p> 
                </div>
            </div>
            <div class="logos">
                <img src="https://companieslogo.com/img/orig/UIS_BIG-d64350be.png?t=1677383940" alt="Logo da Empresa do Analista">
                <img style="height: 100px; width: auto;" src="https://companieslogo.com/img/orig/HEN3.DE-168e26bd.png?t=1593285011" alt="Logo da Empresa do Usuário">
            </div>
            <div class="footer">Esta é uma mensagem automática. Por favor, não responda a este email.</div>
        </div>
    </body>
    </html>
    """.format(nome_analista, nome_usuario, mensagem_elogio, numero_chamado)
    return html_template

def html_para_imagem(html_content, output_path, width, height):
    hti = Html2Image()
    hti.screenshot(html_str=html_content, save_as=output_path, size=(width, height))


def criar_rascunho_outlook(destinatario, assunto, imagem_path, cc=None):
    outlook = win32.Dispatch("Outlook.Application")
    rascunho = outlook.CreateItem(0)  # 0 para email
    rascunho.To = destinatario
    if cc:
        rascunho.CC = ";".join(cc)  # Adiciona os destinatários em cópia
    rascunho.Subject = assunto
    with open(imagem_path, 'rb') as f:
        image_data = f.read()
    image_base64 = base64.b64encode(image_data).decode('utf-8')
    image_html = f'<img src="data:image/png;base64,{image_base64}" alt="Reconhecimento de Excelente Atendimento"/>'
    rascunho.HTMLBody = image_html
    rascunho.Save()

if __name__ == "__main__":
    dados = ler_dados_excel()
    for idx, linha in enumerate(dados):
        nome_analista, nome_usuario, numero_chamado, mensagem_elogio, email_analista, enviado = linha
        cc = ['email1@example.com', 'email2@example.com']  # Adicione os endereços de e-mail CC aqui
        corpo_email = construir_email(nome_analista, nome_usuario, numero_chamado, mensagem_elogio)
        imagem_path = f'email_image_{idx}.png'
        html_para_imagem(corpo_email, imagem_path, 600, 600)
        criar_rascunho_outlook(email_analista, 'Reconhecimento de Excelente Atendimento', imagem_path, cc)
        os.remove(imagem_path)  # Remove o arquivo de imagem após criar o rascunho

