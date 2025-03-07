Attribute VB_Name = "M�dulo2"
Sub AtualizarTextoRegioes()
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' Apresenta��o ativa
    Dim pptSlide As Object ' Slide espec�fico
    Dim pptShape As Object ' Para iterar pelas formas
    Dim excelApp As Object
    Dim excelWB As Object
    Dim excelWS As Object
    Dim Regioes As Object
    Set Regioes = CreateObject("Scripting.Dictionary")

    ' Definir o caminho correto da planilha
    Dim caminhoNovaPlanilha As String
    caminhoNovaPlanilha = "\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresenta��es Padr�o\Pasta1.xlsx"

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
        MsgBox "O PowerPoint n�o est� aberto. Por favor, abra o PowerPoint antes de executar este c�digo.", vbCritical
        excelWB.Close False
        excelApp.Quit
        Exit Sub
    End If
    On Error GoTo 0

    Set pptPres = pptApp.ActivePresentation
    If pptPres Is Nothing Then
        MsgBox "Nenhuma apresenta��o do PowerPoint est� aberta.", vbCritical
        excelWB.Close False
        excelApp.Quit
        Exit Sub
    End If

    ' Definir o slide onde est�o as caixas de texto das regi�es
    Set pptSlide = pptPres.Slides(7)

    ' Varredura da planilha para encontrar as regi�es
    Dim linha As Integer
    linha = 3 ' Come�a na linha 3

    Do While excelWS.Cells(linha, 3).Value <> "" ' Continua enquanto houver nomes de regi�es na coluna C
        Dim nomeRegiao As String
        Dim valorRegiao As String
    
        nomeRegiao = excelWS.Cells(linha, 3).Value ' Nome da regi�o (coluna C)
        valorRegiao = excelWS.Cells(linha, 4).Value ' Valor da regi�o (coluna D)
    
        ' Monta o nome da caixa de texto correspondente no PowerPoint
        Dim nomeCaixa As String
        nomeCaixa = "Caixa" & nomeRegiao ' Ex: "CaixaSul", "CaixaNordeste"...
    
        ' Adiciona ao dicion�rio
        Regioes.Add nomeCaixa, valorRegiao
    
        linha = linha + 1 ' Passa para a pr�xima linha
    Loop

    ' Atualizar os valores das regi�es no PowerPoint
    Dim regiao As Variant
    For Each regiao In Regioes.Keys
        On Error Resume Next
        Set pptShape = pptSlide.Shapes(regiao) ' Nome da forma no PowerPoint deve ser "CaixaSul", "CaixaNordeste", etc.
        If Not pptShape Is Nothing Then
            pptShape.TextFrame.TextRange.Text = Regioes(regiao) ' AGORA APENAS O N�MERO
            Debug.Print "Texto atualizado para regi�o " & Replace(regiao, "Caixa", "") & ": " & Regioes(regiao)
        Else
            Debug.Print "Forma n�o encontrada para regi�o: " & regiao
        End If
        On Error GoTo 0
    Next regiao

    ' Fechar Excel sem salvar altera��es
    excelWB.Close False
    excelApp.Quit

    MsgBox "Atualiza��o das regi�es conclu�da!", vbInformation
End Sub

