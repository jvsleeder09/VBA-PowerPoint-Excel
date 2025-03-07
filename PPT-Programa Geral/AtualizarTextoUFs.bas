Attribute VB_Name = "Módulo1"
Sub AtualizarTextoUFs()
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' Apresentação ativa
    Dim pptSlide As Object ' Slide específico
    Dim pptShape As Object ' Para iterar pelas formas
    Dim valorUF As String
    Dim valorTotalGeral As String
    Dim UF As String
    Dim i As Integer
    Dim excelApp As Object
    Dim excelWB As Object
    Dim excelWS As Object

    ' Lista de UFs
    Dim UFs() As String
    UFs = Split("AC,AL,AM,AP,BA,CE,DF,ES,GO,MA,MG,MS,MT,PA,PB,PE,PI,PR,RJ,RN,RO,RR,RS,SC,SE,SP,TO", ",")

    ' Definir a planilha correta
    Dim caminhoNovaPlanilha As String
    caminhoNovaPlanilha = "\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Pasta1.xlsx"

    ' Iniciar o Excel e abrir a planilha
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
    If excelApp Is Nothing Then
        MsgBox "Erro ao iniciar o Excel.", vbCritical
        Exit Sub
    End If
    excelApp.Visible = False
    Set excelWB = excelApp.Workbooks.Open(caminhoNovaPlanilha, False, True)
    If excelWB Is Nothing Then
        MsgBox "Erro ao abrir a planilha do Excel.", vbCritical
        excelApp.Quit
        Exit Sub
    End If
    Set excelWS = excelWB.Sheets("Planilha1")
    On Error GoTo 0

    ' Conectar ao PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        MsgBox "O PowerPoint não está aberto. Por favor, abra o PowerPoint antes de executar este código.", vbCritical
        excelWB.Close False
        excelApp.Quit
        Exit Sub
    End If
    On Error GoTo 0

    Set pptPres = pptApp.ActivePresentation
    If pptPres Is Nothing Then
        MsgBox "Nenhuma apresentação do PowerPoint está aberta.", vbCritical
        excelWB.Close False
        excelApp.Quit
        Exit Sub
    End If

    ' Definir o slide onde estão as caixas de texto das UFs e o total geral
    Set pptSlide = pptPres.Slides(7)

    ' Atualizar valores para cada UF
    For i = LBound(UFs) To UBound(UFs)
        UF = UFs(i)

        ' Pegar o valor diretamente da planilha
        On Error Resume Next
        valorUF = excelWS.Cells(i + 3, 2).Value ' Ajustado para iniciar na linha 3
        If Err.Number <> 0 Then
            valorUF = "Erro ao buscar dados"
            Debug.Print "Erro ao buscar dados para UF: " & UF
        Else
            Debug.Print "Valor buscado para UF " & UF & ": " & valorUF
        End If
        On Error GoTo 0

        ' Concatenar o nome da UF com o valor
       valorUF = UF & vbCrLf & valorUF ' Adicionando quebra de linha


        ' Buscar a forma e atualizar o texto na caixa de texto correspondente
        On Error Resume Next
        Set pptShape = pptSlide.Shapes("Caixa" & UF)
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = valorUF
            Debug.Print "Texto atualizado para UF " & UF & ": " & valorUF
        Else
            Debug.Print "Forma não encontrada para UF: " & UF
        End If
        On Error GoTo 0
    Next i

    ' Pegar o valor do Total Geral
    On Error Resume Next
    valorTotalGeral = excelWS.Cells(30, 2).Value ' Célula com o total geral (ajuste conforme necessário)
    If Err.Number <> 0 Then
        valorTotalGeral = "Erro ao buscar dados"
        Debug.Print "Erro ao buscar dados para o Total Geral"
    Else
        Debug.Print "Valor buscado para o Total Geral: " & valorTotalGeral
    End If
    On Error GoTo 0

    ' Atualizar a forma do Total Geral
    On Error Resume Next
    Set pptShape = pptSlide.Shapes("CaixaTotalGeral")
    If Not pptShape Is Nothing Then
        pptShape.TextFrame.TextRange.Text = valorTotalGeral
        Debug.Print "Texto atualizado para o Total Geral: " & valorTotalGeral
    Else
        Debug.Print "Forma não encontrada para o Total Geral"
    End If
    On Error GoTo 0

    ' Fechar o Excel
    excelWB.Close False
    excelApp.Quit

    Debug.Print "Atualização concluída."
    MsgBox "Textos das UFs e total geral atualizados com sucesso!", vbInformation
End Sub

