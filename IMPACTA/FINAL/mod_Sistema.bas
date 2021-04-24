Attribute VB_Name = "mod_Sistema"
Public Sub travar_tela()
Dim vobj As Control

For Each vobj In frm_final 'PARA CADA EACH VOBJ IN FRM_FINAL
    If TypeOf vobj Is TextBox Then
    vobj.Locked = True
    vobj.ForeColor = vbBlack
    
    ElseIf TypeOf vobj Is CommandButton Then
    vobj.Enabled = True
    End If
Next



End Sub

Public Sub destravar_tela()
Dim vobj As Control

For Each vobj In frm_final 'PARA CADA EACH VOBJ IN FRM_FINAL
    If TypeOf vobj Is TextBox Then
    vobj.Locked = False
    vobj.ForeColor = vbBlue
    
    ElseIf TypeOf vobj Is CommandButton Then
    vobj.Enabled = False
    End If
Next



End Sub

