Attribute VB_Name = "Module1"
Public Function limpa_tela(Form As Integer)
Dim vobj As Control

If Form = 1 Then
    For Each vobj In frm_CadEdis
        
        If TypeOf vobj Is TextBox Then
            vobj.Text = Empty
         
        ElseIf TypeOf vobj Is ComboBox Then
            vobj.ListIndex = -1
        
        ElseIf TypeOf vobj Is CheckBox Then
            vobj.Value = 0
               
        End If
    Next
    
 
 End If
 


End Function



Public Function Trava_tela(Form As Integer)
Dim vobj As Control

If Form = 1 Then

    For Each vobj In frm_Emails
        If TypeOf vobj Is TextBox Then
            vobj.Enabled = False
            vobj.BackColor = &H80000005
        End If
        
    Next
    
End If


End Function


Public Function Destrava_tela(Form As Integer)
Dim vobj As Control

If Form = 1 Then

    For Each vobj In frm_Emails
        If TypeOf vobj Is TextBox Then
            vobj.Enabled = True
            vobj.BackColor = &HC0FFFF
        End If
        
    Next
    
End If


End Function


Public Function Xlog(xcgc As String, xcliente As String, xtipodoc As String, xhorario As String, xdata As Single, xobs As String)

If xobs = 0 Then
    
    xobs = "ARQUIVO NAO GERADO"
    
Else

    xobs = "ARQUIVO GERADO COM SUCESSO"
    
End If

If xcgc = "" Then

    xcgc = 0
End If


deb_edi.In_Edis_Logs xcgc, xcliente, xtipodoc, xhorario, CDate(xdata), xobs


If deb_edi.rsSel_Grd_Logs.State = 1 Then deb_edi.rsSel_Grd_Logs.Close
    deb_edi.Sel_Grd_Logs Date
    
    frm_verifica.grd_gerados.DataMember = "Sel_Grd_Logs"
    frm_verifica.grd_gerados.Refresh



End Function











