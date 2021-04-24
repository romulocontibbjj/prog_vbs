Attribute VB_Name = "Module1"
Sub cn(Dados As String)
    
    Dim cn As New ADODB.Connection
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\CotyBoys\CotyBoys.mdb;Mode=ReadWrite;Persist Security Info=False"
    
    cn.BeginTrans
    
        cn.Execute (Dados)
        
        If Err.Number <> 0 Then
        
            MsgBox ("Erro") & Err.Description
            cn.RollbackTrans
        Else
            MsgBox ("Inclusão de Cliente efetuado com sucesso"), vbInformation
            cn.CommitTrans
            Exit Sub
        End If
    cn.Close

End Sub


