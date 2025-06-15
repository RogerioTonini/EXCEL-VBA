Sub ListaVarsAmbiente()

   Dim EnvString, Indx, Msg, PathLen    ' Declare variables.
   Indx = 1                             ' Initialize index to 1.
   
   Do
      EnvString = Environ(Indx)                 ' Captura nome da variavel
      Debug.Print EnvString                     ' Imprime a variavel
      If Left(EnvString, 5) = "PATH=" Then      ' Checa as 5 primeiras posições da variavel
         PathLen = Len(Environ("PATH"))         ' Captura o tamanho do conteúdo da variavel
         Msg = "PATH entry = " & Indx & " and length = " & PathLen
      Else
         Indx = Indx + 1    ' Not PATH entry,
      End If    ' so increment.
   Loop Until EnvString = ""
   
   If PathLen > 0 Then
      MsgBox Msg    ' Display message.
   Else
      MsgBox "No PATH environment variable exists."
   End If

End Sub