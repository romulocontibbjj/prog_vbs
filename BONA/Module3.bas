Attribute VB_Name = "Module3"
Public Function ConcatenaString(xvalor As String, xtamanho As Integer, xcaracter As String, xalinhamento As String)

Dim x As Integer
Dim y As Integer

If xalinhamento = "D" Then
    
    ConcatenaString = xvalor & String(xtamanho - Len(xvalor), xcaracter)

ElseIf xalinhamento = "E" Then
       
    ConcatenaString = String(xtamanho - Len(xvalor), xcaracter) & xvalor

ElseIf xalinhamento = "C" Then

    x = (xtamanho - Len(xvalor)) / 2
    
    If (x + x) + Len(xvalor) > xtamanho Then
        y = x - 1
    Else
        y = x
    End If
       
    
    ConcatenaString = String(x, xcaracter) & xvalor & String(y, xcaracter)
     MsgBox Len(ConcatenaString)
     
End If
End Function


Public Function OrdenaLista(XLista As ListBox, xtamanho As Integer, xvalor As String)
Dim x As Integer
Dim y As String
x = xtamanho
xtamanho = xtamanho - 1
Do While xtamanho <> -1

          If xvalor < XLista.List(xtamanho) Then
                y = XLista.List(xtamanho)
                XLista.List(xtamanho) = xvalor
                XLista.List(xtamanho + 1) = y
            
            End If
                    
        
xtamanho = xtamanho - 1

Loop

End Function


Public Function OrdenaListagem(xlist As ListBox)
Dim x As Integer
Dim y As Integer
Dim xaux As String
Dim xlistatamanho As Integer

xlistatamanho = xlist.ListCount - 1

For x = 0 To xlistatamanho
    For y = o To xlistatamanho - 1
        If xlist.List(x) < xlist.List(y) Then
            xaux = xlist.List(y)
            xlist.List(y) = xlist.List(x)
            xlist.List(x) = xaux
        End If
            
    Next
    
Next


End Function



