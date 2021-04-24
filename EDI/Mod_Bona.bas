Attribute VB_Name = "Mod_Bona"
Public Function xbona()

 Dim xfilial As String, xFatura As String, xEmissao As String, xVencto As String, xvalor As String
    Dim xDesconto As String, xnome As String, xEndereco As String, xTelefone As String, xBairro As String
    Dim xCidade As String, xcep As String, xCnpj As String, xBanco As String, xAgencia As String, xNomeBanco As String
    Dim xSituacao As String, xlinha As String
    Dim xrs As Recordset
    Dim xnome1 As String
    
    'busca Faturas não atualizadas
    If de_informa.rsSel_FaturaEDIContabil.State = 1 Then de_informa.rsSel_FaturaEDIContabil.Close
    de_informa.Sel_FaturaEDIContabil
    
    If de_informa.rsSel_FaturaEDIContabil.RecordCount > 0 Then
    
    
    xnome1 = "C:\INFORMA\CONTABIL\M5FAT" & zeros(Day(datahora("DATA")), 2) & _
                                            zeros(Month(datahora("DATA")), 2) & _
                                            zeros(Hour(datahora("HORA")), 2) & _
                                            zeros(Minute(datahora("HORA")), 2) & ".TXT"
        'abre arquivo
        Open xnome1 For Output As #1
                                                
        frm_verifica.xvideolar = xnome1
        
        Do Until de_informa.rsSel_FaturaEDIContabil.EOF
        
            If de_informa.rsSel_FaturaEDIContabil.Fields("status") = "C" And de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "" Then
                de_informa.rsSel_FaturaEDIContabil.MoveNext
            Else
                xfilial = Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 1, 2)
                xFatura = zeros(CDbl(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 3, 6)), 10)
                xEmissao = zeros(Day(de_informa.rsSel_FaturaEDIContabil.Fields("emissao")), 2) & "/"
                xEmissao = xEmissao & zeros(Month(de_informa.rsSel_FaturaEDIContabil.Fields("emissao")), 2) & "/"
                xEmissao = xEmissao & Mid$(Trim$(Str(Year(de_informa.rsSel_FaturaEDIContabil.Fields("emissao")))), 3, 2)
                xVencto = zeros(Day(de_informa.rsSel_FaturaEDIContabil.Fields("vencimento")), 2) & "/"
                xVencto = xVencto & zeros(Month(de_informa.rsSel_FaturaEDIContabil.Fields("vencimento")), 2) & "/"
                xVencto = xVencto & Mid$(Trim$(Str(Year(de_informa.rsSel_FaturaEDIContabil.Fields("vencimento")))), 3, 2)
                xvalor = zeros(Int(de_informa.rsSel_FaturaEDIContabil.Fields("valorfatura") * 100), 14)
                xvalor = Mid$(xvalor, 1, 12) & "." & Mid$(xvalor, 13, 2)
                xDesconto = zeros(Int(de_informa.rsSel_FaturaEDIContabil.Fields("abatimento") * 100), 14)
                xDesconto = Mid$(xDesconto, 1, 12) & "." & Mid$(xDesconto, 13, 2)
                xnome = de_informa.rsSel_FaturaEDIContabil.Fields("cliente_nome") & _
                        String(40 - Len(de_informa.rsSel_FaturaEDIContabil.Fields("cliente_nome")), " ")
                xEndereco = de_informa.rsSel_FaturaEDIContabil.Fields("endcob") & _
                            String(48 - Len(de_informa.rsSel_FaturaEDIContabil.Fields("endcob")), " ")
                xTelefone = de_informa.rsSel_FaturaEDIContabil.Fields("telefonecob") & _
                            String(20 - Len(de_informa.rsSel_FaturaEDIContabil.Fields("telefonecob")), " ")
                xBairro = Space(15)
                xCidade = Mid(de_informa.rsSel_FaturaEDIContabil.Fields("cidadecob"), 1, 25) & _
                          String(25 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cidadecob"), 1, 25)), " ")
                xcep = Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 1, 5) & _
                       String(5 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 1, 5)), " ") & "-"
                xcep = xcep & Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 6, 3) & _
                       String(3 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("cepcob"), 6, 3)), " ")
                xCnpj = Format(de_informa.rsSel_FaturaEDIContabil.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
                xBanco = zeros(CDbl(de_informa.rsSel_FaturaEDIContabil.Fields("banco")), 4)
                xAgencia = zeros(CDbl(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("conta"), 1, 4)), 4)
                xNomeBanco = Mid(de_informa.rsSel_FaturaEDIContabil.Fields("banconome"), 1, 10) & _
                             String(10 - Len(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("banconome"), 1, 10)), " ")
        
                'linha registro
                
                If de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "" Then
                
                    xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                             xBairro & xCidade & xcep & xCnpj & xBanco & xAgencia & xNomeBanco & "I"
                
                    Print #1, xlinha
                
                ElseIf de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "A" Then
                
                    If de_informa.rsSel_FaturaEDIContabil.Fields("status") = "C" Then
                
                        xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                                 xBairro & xCidade & xcep & xCnpj & xBanco & xAgencia & xNomeBanco & "E"
                                 
                        Print #1, xlinha
                        
                    Else
                    
                        xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                                 xBairro & xCidade & xcep & xCnpj & xBanco & xAgencia & xNomeBanco & "E"
                                 
                        Print #1, xlinha
                        
                        xlinha = xfilial & xFatura & xEmissao & xVencto & xvalor & xDesconto & xnome & xEndereco & xTelefone & _
                                 xBairro & xCidade & xcep & xCnpj & xBanco & xAgencia & xNomeBanco & "I"
                        
                        Print #1, xlinha
                        
                    End If
                    
                End If
                    
                de_informa.rsSel_FaturaEDIContabil.MoveNext
                
            End If
        Loop
        
        de_informa.rsSel_FaturaEDIContabil.MoveFirst
        
        Do Until de_informa.rsSel_FaturaEDIContabil.EOF
        
            'ATUALIZA EDI GERADO = S
            de_informa.Alt_AtEdiFatura de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura")
            
            de_informa.rsSel_FaturaEDIContabil.MoveNext
            
        Loop
                    
        Close #1
xbona = 1
    
    Else
        xbona = 0
        Exit Function
    End If




End Function
