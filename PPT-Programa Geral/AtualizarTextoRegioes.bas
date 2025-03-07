Attribute VB_Name = "Módulo2"
Sub AtualizarTextoRegioes()
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' Apresentação ativa
    Dim pptSlide As Object ' Slide específico
    Dim pptShape As Object ' Para iterar pelas formas
    Dim excelApp As Object
    Dim excelWB As Object
    Dim excelWS As Object
    Dim Regioes As Object
    Set Regioes = CreateObject("Scripting.Dictionary")

    ' Definir o caminho correto da planilha
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

    ' Definir o slide onde estão as caixas de texto das regiões
    Set pptSlide = pptPres.Slides(7)

    ' Varredura da planilha para encontrar as regiões
    Dim linha As Integer
    linha = 3 ' Começa na linha 3

    Do While excelWS.Cells(linha, 3).Value <> "" ' Continua enquanto houver nomes de regiões na coluna C
        Dim nomeRegiao As String
        Dim valorRegiao As String
    
        nomeRegiao = excelWS.Cells(linha, 3).Value ' Nome da região (coluna C)
        valorRegiao = excelWS.Cells(linha, 4).Value ' Valor da região (coluna D)
    
        ' Monta o nome da caixa de texto correspondente no PowerPoint
        Dim nomeCaixa As String
        nomeCaixa = "Caixa" & nomeRegiao ' Ex: "CaixaSul", "CaixaNordeste"...
    
        ' Adiciona ao dicionário
        Regioes.Add nomeCaixa, valorRegiao
    
        linha = linha + 1 ' Passa para a próxima linha
    Loop

    ' Atualizar os valores das regiões no PowerPoint
    Dim regiao As Variant
    For Each regiao In Regioes.Keys
        On Error Resume Next
        Set pptShape = pptSlide.Shapes(regiao) ' Nome da forma no PowerPoint deve ser "CaixaSul", "CaixaNordeste", etc.
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = Regioes(regiao) ' AGORA APENAS O NÚMERO
            Debug.Print "Texto atualizado para região " & Replace(regiao, "Caixa", "") & ": " & Regioes(regiao)
        Else
            Debug.Print "Forma não encontrada para região: " & regiao
        End If
        On Error GoTo 0
    Next regiao

    ' Fechar Excel sem salvar alterações
    excelWB.Close False
    excelApp.Quit

    MsgBox "Atualização das regiões concluída!", vbInformation
End Sub

