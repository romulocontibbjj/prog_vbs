Attribute VB_Name = "Module1"
Public Function MarcaTextos(xtext As Control)

xtext.SelStart = 0
xtext.SelLength = Len(xtext.Text)

If xtext.Enabled = True Then
    xtext.SetFocus
End If


xtext.BackColor = &HC0FFFF
xtext.ForeColor = &HC00000

End Function

Public Function DesmarcaTextos(xtext As Control)


xtext.BackColor = &HFFFFFF
xtext.ForeColor = &H80000008

End Function
