# Projeto de Reconhecimento de Atendimento

Este projeto lê dados de um arquivo Excel, constrói emails de reconhecimento de excelente atendimento e cria rascunhos desses emails no Outlook, incorporando uma imagem gerada a partir de um conteúdo HTML.

## Dependências

Certifique-se de ter as seguintes bibliotecas instaladas:

- `openpyxl`: Para ler dados de arquivos Excel.
- `html2image`: Para converter conteúdo HTML em imagens.
- `pywin32`: Para criar rascunhos de email no Outlook.

Para instalar essas dependências, você pode usar o pip:

```sh
pip install openpyxl html2image pywin32
```

## Estrutura do Código

### Funções Principais

- **ler_dados_excel**: Lê dados de um arquivo Excel e retorna as linhas onde a coluna F é "Não".
- **construir_email**: Constrói o conteúdo HTML do email usando informações fornecidas.
- **html_para_imagem**: Converte conteúdo HTML em uma imagem e salva em um caminho especificado.
- **criar_rascunho_outlook**: Cria um rascunho de email no Outlook com a imagem gerada incorporada.

### Fluxo de Trabalho

1. **Ler Dados do Excel**: A função `ler_dados_excel` lê um arquivo Excel e filtra as linhas conforme a necessidade.
2. **Construir Email**: A função `construir_email` recebe as informações do analista, usuário, número do chamado e a mensagem de elogio para gerar o conteúdo HTML do email.
3. **Converter HTML em Imagem**: A função `html_para_imagem` transforma o conteúdo HTML em uma imagem.
4. **Criar Rascunho no Outlook**: A função `criar_rascunho_outlook` cria um rascunho de email no Outlook, incorporando a imagem gerada.

## Uso

### Configuração Inicial

1. **Configurar o Caminho do Arquivo Excel**:
   - Atualize o caminho do arquivo Excel na função `ler_dados_excel` se necessário.

2. **Configurar Destinatários CC**:
   - Adicione os endereços de email CC na lista `cc` no bloco `if __name__ == "__main__":`.

### Executar o Script

Para executar o script, basta rodar:

```sh
python enviaElogio.py
```
