Attribute VB_Name = "M�dulo3"
Sub AtualizarTexto()
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' Apresenta��o ativa
    Dim pptSlide As Object ' Slide espec�fico
    Dim pptShape As Object ' Caixa de texto no slide
    Dim excelApp As Object ' Excel Application
    Dim excelWB As Object ' Workbook do Excel
    Dim excelWS As Object ' Worksheet do Excel
    Dim valorExcel As String ' Valor extra�do da c�lula
    Dim caminhoArquivo As String ' Caminho do arquivo Excel

    ' Definir o caminho correto da planilha
    Dim caminhoNovaPlanilha As String
    caminhoNovaPlanilha = "\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresenta��es Padr�o\Pasta1.xlsx"

    ' Iniciar o Excel
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application") ' Tenta usar um Excel j� aberto
    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application") ' Abre um novo Excel, se necess�rio
    End If
    On Error GoTo 0

    If excelApp Is Nothing Then
        MsgBox "Erro ao iniciar o Excel.", vbCritical
        Exit Sub
    End If

    ' Abrir o arquivo do Excel
    On Error Resume Next
    Set excelWB = excelApp.Workbooks.Open(caminhoNovaPlanilha, False, True) ' ReadOnly = True
    If Err.Number <> 0 Then
        MsgBox "Erro ao abrir o arquivo do Excel.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Definir a planilha
    Set excelWS = excelWB.Sheets("Planilha1")

    ' Iniciar o PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application") ' Tenta usar um PowerPoint j� aberto
    If pptApp Is Nothing Then
        MsgBox "Erro: Nenhum PowerPoint aberto.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Pegar a apresenta��o ativa
    Set pptPres = pptApp.ActivePresentation
    If pptPres Is Nothing Then
        MsgBox "Erro: Nenhuma apresenta��o ativa.", vbCritical
        Exit Sub
    End If

    ' Selecionar o primeiro slide
    Set pptSlide = pptPres.Slides(8)

    ' Atualizar as caixas de texto

    ' CaixaTotalGeral
    AtualizarCaixaDeTexto pptSlide, excelWS, "CaixaTotalGeral", "J5"

     ' CAIXAM (com formata��o de porcentagem)
    AtualizarCaixaDeTexto pptSlide, excelWS, "CaixaM", "K4", "#0.0%"

    ' CAIXAF (com formata��o de porcentagem)
    AtualizarCaixaDeTexto pptSlide, excelWS, "CaixaF", "K3", "#0.0%"

    ' Fechar o Excel
    excelWB.Close False
    excelApp.Quit
    Set excelWB = Nothing
    Set excelApp = Nothing

    MsgBox "Textos atualizados com sucesso!", vbInformation
End Sub

' Fun��o auxiliar para atualizar uma caixa de texto
Sub AtualizarCaixaDeTexto(pptSlide As Object, excelWS As Object, nomeShape As String, celulaExcel As String, Optional formato As String)
    Dim pptShape As Object ' Caixa de texto no slide
    Dim valorExcel As Double ' Valor extra�do da c�lula
    Dim textoFormatado As String ' Texto formatado

    ' Pegar o valor da c�lula
    valorExcel = excelWS.Range(celulaExcel).Value

    ' Formatar o valor, se o formato for fornecido
    If formato <> "" Then
        textoFormatado = Format(valorExcel, formato)
    Else
        textoFormatado = valorExcel
    End If

    ' Verificar se a forma existe antes de tentar modificar
    On Error Resume Next
    Set pptShape = pptSlide.Shapes(nomeShape)
    If pptShape Is Nothing Then
        MsgBox "Erro: A forma '" & nomeShape & "' n�o foi encontrada.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Atualizar o texto da forma
    pptShape.TextFrame.TextRange.Text = textoFormatado
End Sub
