Private Sub OcultarAmbiente()

' Objetivo: Ocultar os objetos: 
'	- Barras Scroll Horizontal, Vertical;
'	- Aba(s) Planilha(s) do arquivo;
'	- Linhas de Grade e Cabeçalhos;
'	- Barra de fórmulas;
'	- Alertas em tela,
'	- Mostrar o aplicativo em Tela cheia (somente a planilha).

	With Windows(ThisWorkbook.Name)
		.DisplayHorinzontalScrollBar = False
		.DisplayVerticalScrollBar = False
		.DisplayWorkbookTabs = False
		.DisplayGridlines = False
		.DisplayHeadings = False
	End With
		
	With Application
		.DisplayFullScreen = True
		.DisplayFormulaBar = False
		.DisplayAlerts = False
	End With
End Sub
