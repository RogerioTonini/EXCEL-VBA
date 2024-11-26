Private Workbook_Activate()

	' Chama a Sub OcultarAmbiente
	OcultarAmbiente()
	
	' Chama a função de ZOOM da Planilha
	Range("[Faixa_Planilha_Que_Contem_Menu]").Select
	ActiveWindow.Zoom = True
	
	[CodeNamePlanilha].Range("A1").Select
End Sub