### `PowerPoint-Atualizacao/AtualizarTextoRegioes.bas - Doação`

* **Funcionalidade:** Atualiza os valores das caixas de texto que representam as regiões geográficas no slide 7 do PowerPoint, utilizando dados da planilha "Pasta1.xlsx". O código lê os nomes das regiões e seus respectivos valores da planilha e atualiza as caixas de texto correspondentes no slide.
* **Planilha:** Pasta1.xlsx
* **Slide:** 7
* **Colunas:** C (nomes das regiões), D (valores das regiões) - linhas 3 em diante.
* **Instruções:**
    1.  Clique no botão localizado na região Norte do slide 7 para executar a macro.
    2.  Certifique-se de que a planilha "Pasta1.xlsx" esteja no caminho especificado no código (`\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Pasta1.xlsx`).
    3.  Os nomes das caixas de texto no PowerPoint devem corresponder aos nomes das regiões na planilha, precedidos por "Caixa" (ex: "CaixaSul", "CaixaNordeste").
    4.  O PowerPoint deve estar aberto para que a atualização ocorra.
    5.  A macro exibirá uma mensagem de conclusão após a atualização.
 


### `PowerPoint-Atualizacao/AtualizarTextoUFs.bas - Doação`

* **Funcionalidade:** Atualiza os valores das caixas de texto que representam as Unidades Federativas (UFs) e o total geral no slide 7 do PowerPoint, utilizando dados da planilha "Pasta1.xlsx". O código lê os valores de cada UF e o total geral da planilha e atualiza as caixas de texto correspondentes no slide.
* **Planilha:** Pasta1.xlsx
* **Slide:** 7
* **Colunas:** B (valores das UFs, linhas 3 em diante), B30 (total geral).
* **Instruções:**
    1.  Certifique-se de que a planilha "Pasta1.xlsx" esteja no caminho especificado no código (`\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Pasta1.xlsx`).
    2.  Os nomes das caixas de texto no PowerPoint devem corresponder aos nomes das UFs, precedidos por "Caixa" (ex: "CaixaSP", "CaixaRJ").
    3.  A caixa de texto para o total geral deve ser nomeada como "CaixaTotalGeral".
    4.  O PowerPoint deve estar aberto para que a atualização ocorra.
    5.  A macro exibirá uma mensagem de conclusão após a atualização.
 


### `PowerPoint-Atualizacao/AtualizarTexto.bas - Formação`

* **Funcionalidade:** Atualiza os valores das caixas de texto no slide 8 do PowerPoint, utilizando dados da planilha "Pasta1.xlsx". O código lê os valores das células especificadas na planilha e atualiza as caixas de texto correspondentes no slide.
* **Planilha:** Pasta1.xlsx
* **Slide:** 8
* **Células:** J5 (CaixaTotalGeral), K4 (CaixaM), K3 (CaixaF).
* **Instruções:**
    1.  Certifique-se de que a planilha "Pasta1.xlsx" esteja no caminho especificado no código (`\\nsarq\deidi\DESFAZIMENTO-COFOR\DESIGN (Cat - Predo - Thi)\Apresentações Padrão\Pasta1.xlsx`).
    2.  As caixas de texto no PowerPoint devem ter os nomes "CaixaTotalGeral", "CaixaM" e "CaixaF".
    3.  A "CaixaM" e "CaixaF" serão formatadas como porcentagem com uma casa decimal.
    4.  O PowerPoint deve estar aberto para que a atualização ocorra.
    5.  A macro exibirá uma mensagem de conclusão após a atualização.
