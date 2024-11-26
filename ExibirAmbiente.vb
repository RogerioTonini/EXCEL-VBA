Private Sub ExibirAmbiente()

' Objetivo: Exibir os objetos: 
'	- Barras Scroll Horizontal, Vertical;
'	- Aba(s) Planilha(s) do arquivo;
'	- Linhas de Grade e Cabeçalhos;
'	- Barra de fórmulas;
'	- Alertas em tela,
'	- Mostrar o aplicativo em Tela cheia (somente a planilha).

	With Windows(ThisWorkbook.Name)
		.DisplayHorinzontalScrollBar = True
		.DisplayVerticalScrollBar = True
		.DisplayWorkbookTabs = True
		.DisplayGridlines = True
		.DisplayHeadings = True
	End With
		
	With Application
		.DisplayFullScreen = False
		.DisplayFormulaBar = True
		.DisplayAlerts = True
	End With
End Sub
