Attribute VB_Name = "Módulo4"
Sub AtualizarListasDesfazimento()
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' Apresentação ativa
    Dim pptSlide As Object ' Slide específico
    Dim pptShape As Object ' Para iterar pelas formas
    Dim valorListas As String
    Dim regiao As String
    Dim i As Integer
    Dim excelApp As Object
    Dim excelWB As Object
    Dim excelWS As Object

    ' Lista de Regiões
    Dim regioes() As String
    regioes = Split("Norte,Nordeste,Centro-Oeste,Sudeste,Sul", ",")

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

        ' Pegar o valor das listas diretamente da planilha
        On Error Resume Next
        valorListas = excelWS.Cells(i + 3, 10).Value ' Coluna J (10) para as listas
        If Err.Number <> 0 Then
            valorListas = "Erro ao buscar dados"
            Debug.Print "Erro ao buscar dados para região: " & regiao
        Else
            Debug.Print "Valor buscado para listas da região " & regiao & ": " & valorListas
        End If
        On Error GoTo 0

        ' Atualizar a forma no PowerPoint
        On Error Resume Next
        Select Case regiao
            Case "Norte"
                Set pptShape = pptSlide.Shapes("CaixaNorte")
            Case "Nordeste"
                Set pptShape = pptSlide.Shapes("CaixaNordeste")
            Case "Centro-Oeste"
                Set pptShape = pptSlide.Shapes("CaixaCentro")
            Case "Sudeste"
                Set pptShape = pptSlide.Shapes("CaixaSudeste")
            Case "Sul"
                Set pptShape = pptSlide.Shapes("CaixaSul")
        End Select

        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = regiao & ": " & valorListas & " listas"
            Debug.Print "Texto atualizado para listas da região " & regiao & ": " & regiao & ": " & valorListas & " listas"
        Else
            Debug.Print "Forma não encontrada para listas da região: " & regiao
        End If
        On Error GoTo 0
    Next i

    ' Fechar o Excel
    excelWB.Close False
    excelApp.Quit

    Debug.Print "Atualização concluída."
    MsgBox "Textos das listas das regiões atualizados com sucesso!", vbInformation
End Sub


