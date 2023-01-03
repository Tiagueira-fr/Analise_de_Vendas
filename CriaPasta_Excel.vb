Sub AdicionarPasta()
  Dim fso As FileSystemObject
  Dim pastaExistente As Folder
  Dim novaPasta As Folder
  
  ' Obtém o objeto FileSystemObject
  Set fso = New FileSystemObject
  
  ' Obtém a pasta existente
  Set pastaExistente = fso.GetFolder("C:\PastaExistente")
  
  ' Cria uma nova pasta dentro da pasta existente
  Set novaPasta = pastaExistente.CreateFolder("NovaPasta")
  
  ' Exibe uma mensagem de sucesso
  MsgBox "Pasta criada com sucesso!"
End Sub
