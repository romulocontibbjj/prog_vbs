Attribute VB_Name = "Correios"
Public Function xCorreios(mskPer1Correio As Single, mskPer2Correio As Single)

    Dim xdestinatario As String, xcep As String, xpeso As Currency, xanotacoes As String, xlinha As String, xfile As String
    
    Dim xnf As String
    Dim xvolumes As String
    Dim xCidade As String
    Dim xend As String
    Dim xuf As String
    
    
    
    If de_informa.rsSel_BuscaCTCCorreios.State = 1 Then de_informa.rsSel_BuscaCTCCorreios.Close
    de_informa.Sel_BuscaCTCCorreios CDate(mskPer1Correio), CDate(mskPer2Correio)
    
    If de_informa.rsSel_BuscaCTCCorreios.RecordCount < 1 Then
        xCorreios = 0
        Exit Function
    Else
    
        xfile = "INT" & zeros(Day(datahora("DATA")), 2) & _
                zeros(Month(datahora("DATA")), 2) & _
                Mid$(datahora("HORA"), 1, 2) & Mid$(datahora("HORA"), 4, 2) & ".TXT"
    
        Open "C:\INFORMA\CORREIOS\" & xfile For Output As #1
        
        frm_verifica.xvideolar = "C:\INFORMA\CORREIOS\" & xfile
    
        Do Until de_informa.rsSel_BuscaCTCCorreios.EOF
        
            xlinha = ""
            
            xdestinatario = Trim$(de_informa.rsSel_BuscaCTCCorreios.Fields("dest_nome")) & _
                            String(50 - Len(Trim$(de_informa.rsSel_BuscaCTCCorreios.Fields("dest_nome"))), " ")
            xcep = de_informa.rsSel_BuscaCTCCorreios.Fields("dest_cep")
            If Len(Trim$(xcep)) < 8 Then
                If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                de_informa.Sel_CadCliCGC de_informa.rsSel_BuscaCTCCorreios.Fields("dest_cgc")
                xcep = de_informa.rsSel_CadCliCGC.Fields("cep")
                If Len(Trim$(xcep)) < 8 Then
                    xcep = "        "
                End If
            End If
          
            xpeso = de_informa.rsSel_BuscaCTCCorreios.Fields("peso") * 1000
            xpesoCHAR = Format(xpeso, "#########0")
            xpesoCHAR = String(10 - Len(xpesoCHAR), "0") & xpesoCHAR
            
            xanotacoes = de_informa.rsSel_BuscaCTCCorreios.Fields("numnf") & "-" & _
                         de_informa.rsSel_BuscaCTCCorreios.Fields("serie")
                         
            xanotacoes = xanotacoes & String(20 - Len(Trim$(xanotacoes)), " ")
            
           
            
        'ADICIONA A VOLUMES
        Dim xvolumes1 As Integer
        
        xvolumes1 = de_informa.rsSel_BuscaCTCCorreios.Fields("volumes")
        xvolumes = xvolumes1
        xvolumes = String(10 - Len(xvolumes), "0") & xvolumes
        'xvolumes = Format(xvolumes1, "#########0")
        
        'ADICONA END
        xend = de_informa.rsSel_BuscaCTCCorreios.Fields("dest_end")
        
        If Len(xend) > 50 Then
            
            xend = Mid(xend, 1, 50)
        
        End If
        
        xend = xend & String(50 - Len(xend), " ")
        'MsgBox Len(xend)
        'ADICIONA CIDADE
        
        xCidade = de_informa.rsSel_BuscaCTCCorreios.Fields("cidade_dest")
        
        If Len(xCidade) > 40 Then
             
            xCidade = Mid(xCidade, 1, 40)
        
        End If
        
        xCidade = xCidade & String(40 - Len(xCidade), " ")
        
        'ADICIONA UF
        
        xuf = de_informa.rsSel_BuscaCTCCorreios.Fields("uf_dest")
        
        If Len(xuf) > 3 Then
            
            xuf = Mid(xuf, 1, 3)
        
        End If
        
        xuf = xuf & String(3 - Len(xuf), " ")
        
        
     
            'xlinha = xdestinatario & xcep & xpesoCHAR & xanotacoes
            xlinha = xdestinatario & xcep & xpesoCHAR & xanotacoes & xvolumes & xCidade & xend & xuf
            
            Print #1, xlinha
            
            de_informa.rsSel_BuscaCTCCorreios.MoveNext

        Loop
        
        Close #1
        
        
        
        xCorreios = 1
        
        
    End If



End Function
