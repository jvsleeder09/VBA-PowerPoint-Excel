### `PowerPoint-Atualizacao/AtualizarTextoS.bas`

* **Funcionalidade:** Atualiza os valores das caixas de texto no slide 4 do PowerPoint, utilizando dados da planilha "Pasta1.xlsx". O código lê os valores das células especificadas na planilha e atualiza as caixas de texto correspondentes no slide.
* **Planilha:** Pasta1.xlsx
* **Slide:** 4
* **Células:** B30 (CaixaTotalGeral), J5 (CaixaTotalGeral2).
* **Instruções:**
    1.  Certifique-se de que a planilha "Pasta1.xlsx" esteja no caminho especificado no código (`\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Pasta1.xlsx`).
    2.  As caixas de texto no PowerPoint devem ter os nomes "CaixaTotalGeral" e "CaixaTotalGeral2".
    3.  O PowerPoint deve estar aberto para que a atualização ocorra.
    4.  A macro exibirá uma mensagem de conclusão após a atualização.
 


### `PowerPoint-Atualizacao/AtualizarListasDesfazimento.bas`

* **Funcionalidade:** Atualiza os valores das caixas de texto que representam as listas de desfazimento por região no slide 7 do PowerPoint, utilizando dados da planilha "Desfazimento.xlsx". O código lê os valores das listas de cada região da planilha e atualiza as caixas de texto correspondentes no slide.
* **Planilha:** Desfazimento.xlsx
* **Slide:** 7
* **Colunas:** J (valores das listas, linhas 3 em diante).
* **Instruções:**
    1.  Certifique-se de que a planilha "Desfazimento.xlsx" esteja no caminho especificado no código (`\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Desfazimento.xlsx`).
    2.  As caixas de texto no PowerPoint devem ter os nomes "CaixaNorte", "CaixaNordeste", "CaixaCentro", "CaixaSudeste" e "CaixaSul".
    3.  O PowerPoint deve estar aberto para que a atualização ocorra.
    4.  A macro exibirá uma mensagem de conclusão após a atualização.



 ### `PowerPoint-Atualizacao/AtualizarDesfazimentoPorRegiao.bas`

* **Funcionalidade:** Atualiza os valores das caixas de texto que representam os itens de desfazimento por região (CPU, NOTE, MONITOR, IMPRESSORA, OUTROS) no slide 7 do PowerPoint, utilizando dados da planilha "Desfazimento.xlsx". O código lê os valores de cada item de desfazimento por região da planilha e atualiza as caixas de texto correspondentes no slide.
* **Planilha:** Desfazimento.xlsx
* **Slide:** 7
* **Colunas:** B (CPU), C (NOTE), D (MONITOR), E (IMPRESSORA), F (OUTROS), linhas 3 em diante.
* **Instruções:**
    1.  Certifique-se de que a planilha "Desfazimento.xlsx" esteja no caminho especificado no código (`\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Desfazimento.xlsx`).
    2.  As caixas de texto no PowerPoint devem ter os nomes "CPUCO", "CPUNo", "CPUNo", "CPUSu", "CPUSu" para CPU, "NOTECO", "NOTENo", "NOTENo", "NOTESu", "NOTESu" para NOTE, e assim por diante para MONITOR, IMPRESSORA e OUTROS.
    3.  O PowerPoint deve estar aberto para que a atualização ocorra.
    4.  A macro exibirá uma mensagem de conclusão após a atualização.
