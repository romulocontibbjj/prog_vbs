Attribute VB_Name = "Module1"
Option Explicit

Public Function Xbusca(ByVal xArquivo As String, Optional xDiretorio As Boolean)
    If Trim$(Len(xArquivo)) > 0 Then
        If IsMissing(xDiretorio) Or xDiretorio = False Then
            Xbusca = (Dir$(xArquivo) <> "")
        Else
            Xbusca = (Dir$(xArquivo, vbDirectory) <> "")
        End If
     End If
End Function
