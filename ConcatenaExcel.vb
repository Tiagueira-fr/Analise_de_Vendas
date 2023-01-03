Sub CopiarArquivo()
  Dim fso As FileSystemObject
  Dim pasta As Folder
  Dim arquivoNovo As File
  Dim arquivoExistente As Workbook
  
  ' Obtém o objeto FileSystemObject
  Set fso = New FileSystemObject
  
  ' Obtém a pasta onde o novo arquivo foi criado
  Set pasta = fso.GetFolder("C:\Pasta")
  
  ' Obtém o novo arquivo criado na pasta
  Set arquivoNovo = pasta.Files(1)
  
  ' Abre o arquivo existente
  Set arquivoExistente = Workbooks.Open("C:\ArquivoExistente.xlsx")
  
  ' Copia todas as folhas do arquivo existente para o novo arquivo
  arquivoExistente.Sheets.Copy After:=arquivoNovo.Workbooks(1).Sheets(1)
  
  ' Fecha o arquivo existente
  arquivoExistente.Close
  
  ' Exibe uma mensagem de sucesso
  MsgBox "Arquivo copiado com sucesso!"
End Sub
