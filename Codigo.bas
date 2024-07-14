Sub ObterValorGoogleFinancas()
    Dim ie As Object
    Dim html As Object
    Dim link As String
    Dim searchResult As Object
    Dim valor As String
    Dim i As Integer
    Dim ws As Worksheet

    ' Inicializa o Internet Explorer invisível
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False

    ' Define a planilha ativa
    Set ws = ActiveSheet

    ' Loop através das células na coluna C, começando da linha 2
    i = 2
    Do While ws.Range("C" & i).Value <> ""
        ' Obtem o link da célula na coluna C
        link = ws.Range("C" & i).Value

        ' Navega para o link
        ie.navigate link

        ' Aguarda a página carregar
        Do While ie.Busy Or ie.readyState <> 4
            DoEvents
        Loop

        ' Obtem o HTML da página
        Set html = ie.document

        ' Verifica se o elemento com classe="P6K39c" está presente
        On Error Resume Next
        Set searchResult = html.getElementsByClassName("P6K39c")(0)
        On Error GoTo 0
        If searchResult Is Nothing Then
            ws.Range("E" & i).Value = "Elemento não encontrado"
        Else
            valor = searchResult.innerText
            ws.Range("E" & i).Value = valor
        End If

        ' Avança para a próxima linha
        i = i + 1
    Loop

    ' Fecha o Internet Explorer
    ie.Quit

    ' Limpa as variáveis
    Set ie = Nothing
    Set html = Nothing
    Set searchResult = Nothing
End Sub