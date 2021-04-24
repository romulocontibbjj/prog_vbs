Attribute VB_Name = "Module1"
Public Function xsepara(xnome As String)
Dim xqtd As Integer
Dim x As Integer
Dim xaux As String

xqtd = Len(xnome)

For x = 1 To xqtd

    If Mid(xnome, x, 1) = "-" Then
        
        xaux = Mid(xnome, 1, x - 1)
        xsepara = xaux
        Exit For
    
    End If



Next




End Function

Public Function Xbusca(ByVal xArquivo As String, Optional xDiretorio As Boolean)

If Trim$(Len(xArquivo)) > 0 Then
    
    If IsMissing(xDiretorio) Or xDiretorio = False Then
        
        Xbusca = (Dir$(xArquivo) <> "")
        
    Else
        
        Xbusca = (Dir$(xArquivo, vbDirectory) <> "")
        
    End If
    
        
 End If
 


End Function

