Attribute VB_Name = "Cobranca"
Public Function DOCCOB(xarquivo As String, txtCgcDocCob As String)

Dim xrs As Recordset, xDataAgora As Date, xHoraAgora As Variant
    Dim xValorTot As Currency, xlinha As String
    
    
        If de_informa.rsSel_EDICobCnpj.State = 1 Then de_informa.rsSel_EDICobCnpj.Close
            de_informa.Sel_EDICobCnpj txtCgcDocCob & "%"
            
            Set xrs = de_informa.rsSel_EDICobCnpj
        
        
            If xrs.RecordCount < 1 Then
                MsgBox "Não Há Dados Para Esta Seleção !", vbInformation
                
                DOCCOB = 0
                
                Exit Function
            End If
    
        
    
        xDataAgora = datahora("data")
        xHoraAgora = datahora("hora")
    
    'abre arquivo
        Open xarquivo For Output As #1
        
                                        
    xlinha = "000INTEC CARGO                        " & Mid$(Trim$(xrs.Fields("cliente_nome")), 1, 35) & _
             String(35 - Len(Mid$(Trim$(xrs.Fields("cliente_nome")), 1, 35)), " ") & zeros(Day(xDataAgora), 2) & _
             zeros(Month(xDataAgora), 2) & Mid$(Trim$(Year(xDataAgora)), 3, 2) & zeros(Hour(xHoraAgora), 2) & _
             zeros(Minute(xHoraAgora), 2) & "COB" & zeros(Day(xDataAgora), 2) & zeros(Month(xDataAgora), 2) & _
             zeros(Hour(xHoraAgora), 2) & zeros(Minute(xHoraAgora), 2) & "0" & Space(75)
            
    Print #1, xlinha
    
    xlinha = "350COBRA" & zeros(Day(xDataAgora), 2) & zeros(Month(xDataAgora), 2) & _
             zeros(Hour(xHoraAgora), 2) & zeros(Minute(xHoraAgora), 2) & "0" & Space(153)
    
    Print #1, xlinha
    
    If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
    de_informa.Sel_CadFilial Mid$(xrs.Fields("filialfatura"), 1, 2)
    
    xlinha = "351" & de_informa.rsSel_CadFilial.Fields("cgc") & "INTEC INTEGRACAO NACIONAL DE TRANSPORTES" & Space(113)
    
    Print #1, xlinha
    
    Do Until xrs.EOF
    
        xlinha = "352" & Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10) & _
                 String(10 - Len(Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10)), " ") & "0" & "U  " & _
                 zeros2(Mid$(xrs.Fields("filialfatura"), 3, 6), 10) & _
                 zeros(Day(xrs.Fields("emissao")), 2) & zeros(Month(xrs.Fields("emissao")), 2) & Trim$(Year(xrs.Fields("emissao"))) & _
                 zeros(Day(xrs.Fields("vencimento")), 2) & zeros(Month(xrs.Fields("vencimento")), 2) & Trim$(Year(xrs.Fields("vencimento"))) & _
                 zeros2(SoNumeros(Format(xrs.Fields("valorfatura"), "#########0.00")), 15) & "   " & _
                 zeros2(SoNumeros(Format(xrs.Fields("descicms"), "#########0.00")), 15) & _
                 "000000000000000" & "00000000" & zeros2(SoNumeros(Format(xrs.Fields("abatimento"), "#########0.00")), 15) & _
                 Trim$(xrs.Fields("banconome")) & String(35 - Len(Trim$(xrs.Fields("banconome"))), " ") & _
                 zeros2(Mid$(xrs.Fields("conta"), 1, InStr(1, xrs.Fields("conta"), ".") - 1), 4) & " " & _
                 zeros2(Mid$(xrs.Fields("conta"), InStr(1, xrs.Fields("conta"), ".") + 1, Abs(InStr(1, xrs.Fields("conta"), "-") - InStr(1, xrs.Fields("conta"), "."))), 10) & _
                 "  " & "I" & Space(3)
                 
        Print #1, xlinha
        
        If de_informa.rsSel_EDICobCTCs.State = 1 Then de_informa.rsSel_EDICobCTCs.Close
        de_informa.Sel_EDICobCTCs xrs.Fields("filialfatura")
        
        Do Until de_informa.rsSel_EDICobCTCs.EOF
        
            If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
            de_informa.Sel_CadFilial Mid$(de_informa.rsSel_EDICobCTCs.Fields("filialctc"), 1, 2)
            
            xlinha = "353" & Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10) & _
                     String(10 - Len(Mid$(Trim$(de_informa.rsSel_CadFilial.Fields("nomefilial")), 1, 10)), " ") & _
                     Trim$(de_informa.rsSel_CadFilial.Fields("seriectc")) & _
                     String(5 - Len(Trim$(de_informa.rsSel_CadFilial.Fields("seriectc"))), " ") & _
                     zeros2(Mid$(de_informa.rsSel_EDICobCTCs.Fields("filialctc"), 3), 12) & Space(140)

            Print #1, xlinha
            
            If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
            de_informa.Sel_NFsdoCTC de_informa.rsSel_EDICobCTCs.Fields("filialctc")
            
            Do Until de_informa.rsSel_NFsdoCTC.EOF
            
                xlinha = "354" & de_informa.rsSel_NFsdoCTC.Fields("serie") & String(3 - Len(de_informa.rsSel_NFsdoCTC.Fields("serie")), " ") & _
                         zeros2(de_informa.rsSel_NFsdoCTC.Fields("numnf"), 8) & String(30, "0") & _
                         zeros2(de_informa.rsSel_NFsdoCTC.Fields("cliente_cgc"), 14) & Space(112)
            
                Print #1, xlinha
                
                de_informa.rsSel_NFsdoCTC.MoveNext
                
            Loop
            
            de_informa.rsSel_EDICobCTCs.MoveNext
            
        Loop
        
        xValorTot = xValorTot + xrs.Fields("valorfatura")
        
        xrs.MoveNext
        
    Loop
    
    xlinha = "355" & zeros(xrs.RecordCount, 4) & zeros2(SoNumeros(Format(xValorTot, "#########0.00")), 15) & Space(148)
    
    Print #1, xlinha
    
    Close #1
    
    xrs.MoveFirst
    
    Do Until xrs.EOF
        de_informa.Alt_EDICobATSim xrs.Fields("filialfatura")
        xrs.MoveNext
    Loop
    
    
    
End Function
