Attribute VB_Name = "NFMEDLEY"
Public Function ConembMedley(xarquivo As String)

frm_verifica.MousePointer = 11
    
    If de_informa.rsSel_EDICtcsMedley.State = 1 Then de_informa.rsSel_EDICtcsMedley.Close
    de_informa.Sel_EDICtcsMedley "50929710000179"

    If de_informa.rsSel_EDICtcsMedley.RecordCount > 0 Then
    
        'abre o arquivo
        
        Open xarquivo For Output As #1
        
        Do Until de_informa.rsSel_EDICtcsMedley.EOF
        
            'verifica o consignatário para confirmar que é da operação Bomi/Intec
            If Not (Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "50929710" Or _
                    Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "04019475" Or _
                    Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "52134798" Or _
                    Mid$(de_informa.rsSel_EDICtcsMedley.Fields("respons_cgc"), 1, 8) = "02426290") Then
                de_informa.rsSel_EDICtcsMedley.MoveNext
            Else
        
                'tratamento da série da NF
            
                If de_informa.rsSel_EDICtcsMedley.Fields("numnfnum") > 200000 Then
                    xserie = "1  "
                Else
                    xserie = "2  "
                End If
                
                xNumNF = zeros(de_informa.rsSel_EDICtcsMedley.Fields("numnfnum"), 6)
                
                If Not IsNull(de_informa.rsSel_EDICtcsMedley.Fields("pesonf")) Then
                    xpeso = zeros(de_informa.rsSel_EDICtcsMedley.Fields("pesonf") * 1000, 13)
                Else
                    xpeso = "0000000000000"
                End If
                
                xctc = "BOM" & Mid$(zeros(Val(de_informa.rsSel_EDICtcsMedley.Fields("ctc")), 8), 3, 6)
            
                If Not IsNull(de_informa.rsSel_EDICtcsMedley.Fields("volumesnf")) Then
                    xvolumes = zeros(de_informa.rsSel_EDICtcsMedley.Fields("volumesnf"), 6)
                Else
                    xvolumes = "000000"
                End If
                

            
                'data de embarque
                
                xdata = ""
                xdata = Trim$(Str(Year(de_informa.rsSel_EDICtcsMedley.Fields("data"))))
                xdata = xdata & zeros(Month(de_informa.rsSel_EDICtcsMedley.Fields("data")), 2)
                xdata = xdata & zeros(Day(de_informa.rsSel_EDICtcsMedley.Fields("data")), 2)
                
                'linha registro
                
                xlinha = "50" & "0001" & xserie & xNumNF & xpeso & xctc & xvolumes & xdata
                
                Print #1, xlinha
                
                'ATUALIZA EDI GERADO = S
                
                de_informa.alt_EDICtcMedleySim de_informa.rsSel_EDICtcsMedley.Fields("filialctc")
                
                de_informa.rsSel_EDICtcsMedley.MoveNext
                    
            End If
            DoEvents
        Loop
        Close #1
        
       ConembMedley = 1
       
    
    Else
        ConembMedley = 0
    End If
    
    frm_verifica.MousePointer = 0
    
End Function
