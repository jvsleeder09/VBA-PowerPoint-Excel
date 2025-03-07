Attribute VB_Name = "Módulo1"
Sub AtualizarDesfazimentoPorRegiao()
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' Apresentação ativa
    Dim pptSlide As Object ' Slide específico
    Dim pptShape As Object ' Para iterar pelas formas
    Dim valorCPU As String, valorNote As String, valorMonitor As String, valorImpressora As String, valorOutros As String
    Dim regiao As String
    Dim i As Integer
    Dim excelApp As Object
    Dim excelWB As Object
    Dim excelWS As Object

    ' Lista de Regiões
    Dim regioes() As String
    regioes = Split("Centro-Oeste,Nordeste,Norte,Sudeste,Sul", ",")

    ' Definir a planilha correta
    Dim caminhoNovaPlanilha As String
    caminhoNovaPlanilha = "\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Desfazimento.xlsx"

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

    ' Definir o slide onde estão as caixas de texto das regiões
    Set pptSlide = pptPres.Slides(7)

    ' Atualizar valores para cada região
    For i = LBound(regioes) To UBound(regioes)
        regiao = regioes(i)

        ' Pegar os valores diretamente da planilha
        On Error Resume Next
        valorCPU = excelWS.Cells(i + 3, 2).Value
        valorNote = excelWS.Cells(i + 3, 3).Value
        valorMonitor = excelWS.Cells(i + 3, 4).Value
        valorImpressora = excelWS.Cells(i + 3, 5).Value
        valorOutros = excelWS.Cells(i + 3, 6).Value
        If Err.Number <> 0 Then
            valorCPU = "Erro"
            valorNote = "Erro"
            valorMonitor = "Erro"
            valorImpressora = "Erro"
            valorOutros = "Erro"
            Debug.Print "Erro ao buscar dados para região: " & regiao
        Else
            Debug.Print "Valores buscados para região " & regiao & ": CPU=" & valorCPU & ", Note=" & valorNote & ", Monitor=" & valorMonitor & ", Impressora=" & valorImpressora & ", Outros=" & valorOutros
        End If
        On Error GoTo 0

        ' Atualizar as formas no PowerPoint
        On Error Resume Next
        Set pptShape = pptSlide.Shapes("CPU" & Left(regiao, 2)) ' Usando os dois primeiros caracteres da região
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = valorCPU
            Debug.Print "Texto atualizado para CPU " & regiao & ": " & valorCPU
        Else
            Debug.Print "Forma não encontrada para CPU: " & regiao
        End If

        Set pptShape = pptSlide.Shapes("NOTE" & Left(regiao, 2))
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = valorNote
            Debug.Print "Texto atualizado para NOTE " & regiao & ": " & valorNote
        Else
            Debug.Print "Forma não encontrada para NOTE: " & regiao
        End If

        Set pptShape = pptSlide.Shapes("MONITOR" & Left(regiao, 2))
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = valorMonitor
            Debug.Print "Texto atualizado para MONITOR " & regiao & ": " & valorMonitor
        Else
            Debug.Print "Forma não encontrada para MONITOR: " & regiao
        End If

        Set pptShape = pptSlide.Shapes("IMPRESSORA" & Left(regiao, 2))
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = valorImpressora
            Debug.Print "Texto atualizado para IMPRESSORA " & regiao & ": " & valorImpressora
        Else
            Debug.Print "Forma não encontrada para IMPRESSORA: " & regiao
        End If

        Set pptShape = pptSlide.Shapes("OUTROS" & Left(regiao, 2))
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = valorOutros
            Debug.Print "Texto atualizado para OUTROS " & regiao & ": " & valorOutros
        Else
            Debug.Print "Forma não encontrada para OUTROS: " & regiao
        End If
        On Error GoTo 0
    Next i

    ' Fechar o Excel
    excelWB.Close False
    excelApp.Quit

    Debug.Print "Atualização concluída."
    MsgBox "Textos das regiões atualizados com sucesso!", vbInformation
End Sub
