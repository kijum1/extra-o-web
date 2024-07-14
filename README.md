ObterValorGoogleFinancas
Este script VBA acessa o Google Finanças para obter valores financeiros de links especificados em uma planilha do Excel. Ele usa o Internet Explorer para navegar até os links e extrair os valores, salvando-os na mesma planilha.

Funcionalidade
O script faz o seguinte:

Inicializa o Internet Explorer de forma invisível.
Define a planilha ativa.
Percorre as células na coluna C a partir da linha 2.
Para cada link na coluna C, o script:
Navega para o link.
Aguarda a página carregar.
Extrai o valor do elemento com a classe "P6K39c".
Salva o valor na coluna E.
Repete o processo até que uma célula vazia seja encontrada na coluna C.
Fecha o Internet Explorer e limpa as variáveis.
Modo de Usar
Preparação da Planilha
Insira os links do Google Finanças na coluna C, começando da célula C2.
Certifique-se de que a coluna E está disponível para os valores extraídos.
Executando o Script
Abra o Excel e pressione Alt + F11 para abrir o Editor do VBA.
Insira um novo módulo:
Vá para Inserir > Módulo.
Cole o código fornecido no módulo.  
