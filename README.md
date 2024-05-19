# Survey Email

Este é um projeto que envolve o envio de e-mails para pesquisas.

## Descrição

Este projeto consiste em um script Python para automatizar o envio de e-mails de pesquisa. O script lê dados de uma planilha Excel, seleciona os destinatários com base em determinados critérios e envia e-mails personalizados para esses destinatários. Além disso, os resultados da pesquisa são registrados de volta na planilha para análise posterior.

## Funcionalidades

- Leitura de dados de uma planilha Excel.
- Seleção de destinatários com base em critérios específicos.
- Envio de e-mails personalizados.
- Registro dos resultados da pesquisa de volta na planilha.

## Como usar

1. Clone este repositório para o seu ambiente local:
    ```bash
    git clone https://github.com/GuilhermeSavioRibas/Survey-email.git
    ```
2. Navegue até o diretório do projeto:
    ```bash
    cd Survey-email
    ```
3. Certifique-se de ter as dependências instaladas. Você pode instalar as dependências usando o pip:
    ```bash
    pip install openpyxl html2image pillow pywin32
    ```
4. Execute o script `enviaElogio.py`:
    ```bash
    python enviaElogio.py
    ```
5. Verifique a caixa de saída do seu cliente de e-mail para confirmar o envio dos e-mails.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir problemas (issues) ou enviar pull requests com melhorias, correções de bugs ou novas funcionalidades.

## Licença

Este projeto é licenciado sob a [MIT License](LICENSE).
