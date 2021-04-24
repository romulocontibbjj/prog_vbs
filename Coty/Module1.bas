Attribute VB_Name = "Module1"
Option Explicit

Public Function limpa_tela(Form As Integer)
Dim vobj As Control

If Form = 1 Then
    For Each vobj In FRM_FichadeClientes
        
        If TypeOf vobj Is TextBox Then
            vobj.Text = Empty
         
        ElseIf TypeOf vobj Is ComboBox Then
            vobj.ListIndex = -1
        
        ElseIf TypeOf vobj Is MaskEdBox Then
            vobj.Mask = ""
            vobj.Text = ""
            vobj.Mask = "99/99/9999"
               
        End If
    Next
    FRM_FichadeClientes.txtRazao.SetFocus

ElseIf Form = 2 Then

        For Each vobj In FRM_FichadeMotoqueiros
        
        If TypeOf vobj Is TextBox Then
            vobj.Text = Empty
         
        ElseIf TypeOf vobj Is ComboBox Then
            vobj.ListIndex = -1
        
        ElseIf TypeOf vobj Is MaskEdBox Then
            vobj.Mask = ""
            vobj.Text = ""
            vobj.Mask = "99/99/9999"
                    
        End If
    Next
ElseIf Form = 3 Then

        For Each vobj In FRM_FichadeMotos
        
        If TypeOf vobj Is TextBox Then
            vobj.Text = Empty
         
        ElseIf TypeOf vobj Is ComboBox Then
            vobj.ListIndex = -1
        
        ElseIf TypeOf vobj Is MaskEdBox Then
            vobj.Mask = ""
            vobj.Text = ""
            vobj.Mask = "99/99/9999"
                    
        End If
    Next

End If


End Function


Public Function branco_separa(xstring As String)

Dim x As Integer

    For x = 1 To Len(xstring)
        branco_separa = Mid(xstring, x, 1)
            If branco_separa = " " Then
                branco_separa = Mid(xstring, 1, x - 1)
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
