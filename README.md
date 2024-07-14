# ObterValorGoogleFinancas

Este script VBA acessa o Google Finanças para obter valores financeiros de links especificados em uma planilha do Excel. Ele usa o Internet Explorer para navegar até os links e extrair os valores, salvando-os na mesma planilha.

## Funcionalidade

O script faz o seguinte:

1. Inicializa o Internet Explorer de forma invisível.
2. Define a planilha ativa.
3. Percorre as células na coluna C a partir da linha 2.
4. Para cada link na coluna C, o script:
   - Navega para o link.
   - Aguarda a página carregar.
   - Extrai o valor do elemento com a classe `P6K39c`.
   - Salva o valor na coluna E.
5. Repete o processo até que uma célula vazia seja encontrada na coluna C.
6. Fecha o Internet Explorer e limpa as variáveis.

## Modo de Usar

### Preparação da Planilha

1. Insira os links do Google Finanças na coluna C, começando da célula C2.
2. Certifique-se de que a coluna E está disponível para os valores extraídos.

### Executando o Script

1. Abra o Excel e pressione `Alt + F11` para abrir o Editor do VBA.
2. Insira um novo módulo:
   - Vá para `Inserir > Módulo`.
3. Cole o código fornecido no módulo.
