import xlwings as xw
from html2image import Html2Image
import base64
import win32com.client as win32
import os

def ler_dados_excel():
    arquivo_excel = r'C:\workspace\Survey-email\Pasta1.xlsm'
    wb = xw.Book(arquivo_excel)
    planilha = wb.sheets[0]
    dados = []
    # Lê os dados da planilha
    for row in range(2, planilha.range('A' + str(planilha.cells.last_cell.row)).end('up').row + 1):
        if planilha.range(f'F{row}').value == "Não":  # Verifica se a coluna enviou_email (F) é "Não"
            dados.append(row)
    return dados, wb, planilha

def construir_email(nome_analista, nome_usuario, numero_chamado, mensagem_elogio, logo_usuario):
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
                height: 50px; /* Defina a altura fixa */
                width: 100px; /* Defina a largura fixa */
                object-fit: contain; /* Mantenha a proporção da imagem */
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
                <img src="{}" alt="Logo da Empresa do Usuário">
            </div>
            <div class="footer"></div>
        </div>
    </body>
    </html>
    """.format(nome_analista, nome_usuario, mensagem_elogio, numero_chamado, logo_usuario)
    return html_template

def html_para_imagem(html_content, output_path, width, height):
    hti = Html2Image()
    hti.screenshot(html_str=html_content, save_as=output_path, size=(width, height))

def criar_rascunho_outlook(destinatario, nome_cliente, assunto_base, imagem_path, cc=None):
    outlook = win32.Dispatch("Outlook.Application")
    rascunho = outlook.CreateItem(0)  # 0 para email
    rascunho.To = destinatario
    if cc:
        rascunho.CC = ";".join(cc)  # Adiciona os destinatários em cópia
    
    # Modificando o assunto para incluir o nome do cliente
    assunto = f"[{nome_cliente}] {assunto_base}"
    rascunho.Subject = assunto
    
    # Carregar o link da logo da empresa do usuário com base no nome do cliente
    logo_usuario = logos_clientes.get(nome_cliente, "Link para o logo padrão")
    
    if os.path.exists(imagem_path):
        with open(imagem_path, 'rb') as f:
            image_data = f.read()
        image_base64 = base64.b64encode(image_data).decode('utf-8')
        image_html = f'<img src="data:image/png;base64,{image_base64}" alt="Reconhecimento de Excelente Atendimento"/>'
        rascunho.HTMLBody = image_html
        rascunho.Display()  # Abre o rascunho no Outlook
    else:
        print(f"Erro: O arquivo de imagem {imagem_path} não foi encontrado.")


if __name__ == "__main__":
    # Definindo os links das logos das empresas dos usuários
    logos_clientes = {
        "Flowserve": "https://companieslogo.com/img/orig/FLS-2ff8c8f5.png?t=1683790943",
        "UUS": "https://companieslogo.com/img/orig/UIS_BIG-d64350be.png?t=1677383940",
        "Cteep": "https://www.isacteep.com.br/Arquivos/Imagens/logo-cteep-face.jpg",
        "Heineken": "https://companieslogo.com/img/orig/HEIA.AS_BIG-46c1e364.png?t=1665028384",
        "Alpek": "https://alpekpolyester.com.br/wp-content/uploads/2019/01/alpek_polyesterTM-Logo.png",
        "Unilever": "https://companieslogo.com/img/orig/UL_BIG-593a9828.png?t=1633508892",
        "Henkel": "https://companieslogo.com/img/orig/HEN3.DE-168e26bd.png?t=1593285011",
        # Adicione os links para os outros clientes conforme necessário

    }
    
    dados, wb, planilha = ler_dados_excel()
    for idx in dados:
        nome_analista = planilha.range(f'A{idx}').value
        nome_usuario = planilha.range(f'B{idx}').value
        numero_chamado = planilha.range(f'C{idx}').value
        mensagem_elogio = planilha.range(f'D{idx}').value
        email_analista = planilha.range(f'E{idx}').value
        nome_cliente = planilha.range(f'H{idx}').value

        cc = ['email1@example.com', 'email2@example.com']  # Adicione os endereços de e-mail CC aqui
        # alterar o campo cc para puxar do excel de acordo com a empresa e grupo.
        
        corpo_email = construir_email(nome_analista, nome_usuario, numero_chamado, mensagem_elogio, logos_clientes.get(nome_cliente, "Link padrão da logo do usuário"))
        imagem_path = f'email_image_{idx}.png'
        
        # Criar a imagem a partir do HTML
        html_para_imagem(corpo_email, imagem_path, 600, 600)
        
        # Verifique se o arquivo de imagem foi criado
        if os.path.exists(imagem_path):
            criar_rascunho_outlook(email_analista, nome_cliente, 'Reconhecimento de Excelente Atendimento', imagem_path, cc)
            # Atualizar a coluna enviou_email (F) para "Sim"
            planilha.range(f'F{idx}').value = "Sim"
            os.remove(imagem_path)  # Remove o arquivo de imagem após criar o rascunho
        else:
            print(f"Erro: A imagem {imagem_path} não foi criada corretamente.")
    
    # Salvar alterações na planilha
    wb.save()
