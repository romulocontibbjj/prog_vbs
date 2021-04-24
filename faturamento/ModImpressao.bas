Attribute VB_Name = "ModImpressao"
'Teste
Public Sub imprime_fat(xFilialFatura As String)
    Dim ximpr_inst As Printer, ximpr_cfg As String, xnumCTC As String, xlinha As String
    Dim xlin As Integer, xcol As Integer, xcont As Integer, xValorExt As String, xFilialImpressao As String

    'busca impressora para este documento
    If Dir("C:\informa.cfg") <> "" Then
        
        Open "C:\informa.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "FAT" Then
                If Trim$(Mid$(xlinha, 5, 2)) = "\\" Then
                    ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Else
                    ximpr_cfg = "LPT1:"
                End If
                Exit Do
            End If
        Loop
        
        If EOF(1) Then
            MsgBox "Não está Configurado a Impressora para Este Documento: CTC " & xFilialCtc
            Close #1
            Exit Sub
        End If
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: CTC " & xFilialCtc
        Exit Sub
        
    End If

    'BUSCA CTC A SER IMPRESSO
    If de_informa.rsSel_Fatura.State = 1 Then de_informa.rsSel_Fatura.Close
    de_informa.Sel_Fatura xFilialFatura
    
    If de_informa.rsSel_Fatura.RecordCount < 1 Then
        MsgBox "Fatura Inexistente !", vbInformation
        Exit Sub
    End If
    

    'xFilialFatura - Variável q controla o tipo de FORMULÁRIO DE FATURA UTILIZADO
    xFilialImpressao = Mid(xFilialFatura, 1, 2)
      
    If de_informa.rsSel_FilialImpressao.State = 1 Then de_informa.rsSel_FilialImpressao.Close
    
    de_informa.Sel_FilialImpressao xFilialImpressao
       
        
    If de_informa.rsSel_FilialImpressao.EOF Then
        MsgBox "Não foi encontrada configuração de formulário para esta FILIAL " & xFilialImpressao
        Close #1
        Exit Sub
    Else
        '*************************Formulário BOMI FARMA
        If de_informa.rsSel_FilialImpressao("formularioFatura") = 2 Then
            
            Open ximpr_cfg For Output As #1
            DoEvents
                
            'inicia a impressão
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                            mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                            Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
            Print #1, String(13 - Len(Trim$(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"))), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"); "      ";
            Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
            Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
            
            Print #1, ""
            Print #1, ""
        
            Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
                
            Print #1, ""
                
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
            Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
            Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
            
            'CGC + FORMATAÇÃO
            If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
            Else
                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
            End If
            
            Print #1, Space(21); xcgc1;
            Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
            
            Print #1, ""
            Print #1, ""
            
            'impreme o valor por extenso
            
            xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
            
            Print #1, Space(21); Mid$(xValorExt, 1, 58)
            Print #1, Space(21); Mid$(xValorExt, 59, 58)
            Print #1, Space(21); Mid$(xValorExt, 117, 58)
            
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1, ""
            Print #1, ""
            'Print #1, ""
            
            Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
            Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
            Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
            Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            If de_informa.rsSel_Fatura.Fields("avulsa") <> "AVULSA" Then
            
                'BUSCA CTC A SER IMPRESSO
                If de_informa.rsSel_FaturaItens.State = 1 Then de_informa.rsSel_FaturaItens.Close
                de_informa.Sel_FaturaItens xFilialFatura
            
                xlin = 1
                xcol = 1
                
                Do Until de_informa.rsSel_FaturaItens.EOF
                
                    'Franklin Modifiquei abaixo ...
                    If de_informa.rsSel_FaturaItens.Fields("tipodoc") = "NFS" Then
                        Print #1, Spc(2); Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); " ";
                    Else
                        Print #1, Spc(1); Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); ' " ";
                    End If
                    'Fim
                    
                    If xcol = 3 Then 'terceira coluna
                        Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")
                        xlin = xlin + 1
                        xcol = 0
                        If xlin = 28 Then
                            
                            For xcont = 1 To 12
                                Print #1, ""
                            Next
                            
                            Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                                            mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                                            Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            
                            Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
                            'Print #1, Space(4); "**" & "      "
                            Print #1, String(13 - Len(Trim$("********")), " "); "********"; "      ";
                            Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
                            Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            'Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
                                
                            Print #1, ""
                            'Print #1, ""
                            
                                                    
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
                            Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
                            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
                            Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
                            
                            'CGC + FORMATAÇÃO
                            If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
                                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
                            Else
                                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
                            End If
                            
                            Print #1, Space(21); xcgc1;
                            Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
                            
                            Print #1,
                            Print #1,
                            Print #1, Space(21); "**"
                            'impreme o valor por extenso
                            
                            'xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
                            
                            'Print #1, Space(21); Mid$(xValorExt, 1, 58)
                            'Print #1, Space(21); Mid$(xValorExt, 59, 58)
                            'Print #1, Space(21); Mid$(xValorExt, 117, 58)
                            
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            
                            
                            
                            Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
                            Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
                            Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
                            Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
                            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                                                
                            xlin = 1
                        
                        End If
                    Else
                        Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00"); " ";
                    End If
                    
                    xcol = xcol + 1
                    
                    de_informa.rsSel_FaturaItens.MoveNext
                                
                Loop
                
                For xcont = 1 To 31 - xlin
                    Print #1, ""
                Next
                
            Else
                For xcont = 1 To 31
                    Print #1, ""
                Next
            End If
            
        
            
            'Print #1, ""
            'Print #1, ""
            
            Print #1, Space(9); String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "###,##0.00");
            Print #1, Space(12); String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00");
            Print #1, Space(25); String(14 - Len(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00")
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Close #1


        '************************* Formulário BOMI BRASIL
        ElseIf de_informa.rsSel_FilialImpressao("formularioFatura") = 3 Then
            
            Open ximpr_cfg For Output As #1
            DoEvents
                
        '    If ximpr_cfg = "LPT1" Then
        '        Open ximpr_cfg For Output As #1
        '        DoEvents
        '    Else
        '        For Each ximpr_inst In Printers
        '            If ximpr_inst.DeviceName = ximpr_cfg Then
        '                Open ximpr_cfg For Output As #1
        '                DoEvents
        '                Exit For
        '            End If
        '        Next
        '    End If
        
            'inicia a impressão
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                            mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                            Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
            Print #1, String(13 - Len(Trim$(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"))), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"); "      ";
            Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
            Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
            
            Print #1, ""
            Print #1, ""
        
            Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
                
            Print #1, ""
                
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
            Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
            Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
            
            'CGC + FORMATAÇÃO
            If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
            Else
                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
            End If
            
            Print #1, Space(21); xcgc1;
            Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
            
            Print #1, ""
            Print #1, ""
            
            'imprime o valor por extenso
            
            xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
            
            Print #1, Space(21); Mid$(xValorExt, 1, 58)
            Print #1, Space(21); Mid$(xValorExt, 59, 58)
            Print #1, Space(21); Mid$(xValorExt, 117, 58)
            
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1, ""
            Print #1, ""
            'Print #1, ""
            
            'Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
            'Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
            'Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
            'Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
            'Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
            
            'Print #1, ""
            'Print #1, ""
            'Print #1, ""
            'Print #1, ""
            'Print #1, ""
            
            
            'Print #1, ""
            'Print #1, ""
            'Print #1, ""
            'Print #1, ""
            'Print #1, ""
            
            If de_informa.rsSel_Fatura.Fields("avulsa") <> "AVULSA" Then
            
                'BUSCA CTC A SER IMPRESSO
                If de_informa.rsSel_FaturaItens.State = 1 Then de_informa.rsSel_FaturaItens.Close
                de_informa.Sel_FaturaItens xFilialFatura
            
                xlin = 1
                xcol = 1
                
                Do Until de_informa.rsSel_FaturaItens.EOF
                
                    'Franklin Modifiquei abaixo ...
                    If de_informa.rsSel_FaturaItens.Fields("tipodoc") = "NFS" Then
                        Print #1, Spc(2); Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); " ";
                    Else
                        Print #1, Spc(1); Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); ' " ";
                    End If
                    'Fim
                    
                    If xcol = 3 Then 'terceira coluna
                        Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")
                        xlin = xlin + 1
                        xcol = 0
                        If xlin = 35 Then
                            
                            For xcont = 1 To 12
                                Print #1, ""
                            Next
                            
                            Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                                            mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                                            Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            
                            Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
                            'Print #1, Space(4); "**" & "      "
                            Print #1, String(13 - Len(Trim$("********")), " "); "********"; "      ";
                            Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
                            Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            'Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
                                
                            Print #1, ""
                            'Print #1, ""
                            
                                                    
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
                            Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
                            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
                            Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
                            
                            'CGC + FORMATAÇÃO
                            If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
                                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
                            Else
                                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
                            End If
                            
                            Print #1, Space(21); xcgc1;
                            Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
                            
                            Print #1,
                            Print #1,
                            Print #1, Space(21); "**"
                            'impreme o valor por extenso
                            
                            'xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
                            
                            'Print #1, Space(21); Mid$(xValorExt, 1, 58)
                            'Print #1, Space(21); Mid$(xValorExt, 59, 58)
                            'Print #1, Space(21); Mid$(xValorExt, 117, 58)
                            
                            Print #1,
                            Print #1,
                            Print #1,
                            
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            
                            'Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
                            'Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
                            'Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
                            'Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
                            'Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                                                
                            xlin = 1
                        
                        End If
                    Else
                        Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00"); " ";
                    End If
                    
                    xcol = xcol + 1
                    
                    de_informa.rsSel_FaturaItens.MoveNext
                                
                Loop
                
                For xcont = 1 To 37 - xlin
                    Print #1, ""
                Next
                
            Else
                For xcont = 1 To 32
                    Print #1, ""
                Next
            End If
            
        
            
            'Print #1, ""
            'Print #1, ""
            
            Print #1, Space(9); String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "###,##0.00");
            Print #1, Space(12); String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00");
            Print #1, Space(25); String(14 - Len(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00")
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Close #1


        '*************************Formulário Intec
        Else
            Open ximpr_cfg For Output As #1
            DoEvents
                
        '    If ximpr_cfg = "LPT1" Then
        '        Open ximpr_cfg For Output As #1
        '        DoEvents
        '    Else
        '        For Each ximpr_inst In Printers
        '            If ximpr_inst.DeviceName = ximpr_cfg Then
        '                Open ximpr_cfg For Output As #1
        '                DoEvents
        '                Exit For
        '            End If
        '        Next
        '    End If
        
            'inicia a impressão
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                            mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                            Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
            Print #1, String(13 - Len(Trim$(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"))), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"); "      ";
            Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
            Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
            
            Print #1, ""
            Print #1, ""
        
            Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
                
            Print #1, ""
                
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
            Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
            Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
            
            'CGC + FORMATAÇÃO
            If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
            Else
                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
            End If
            
            Print #1, Space(21); xcgc1;
            Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
            
            Print #1, ""
            
            'impreme o valor por extenso
            
            xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
            
            Print #1, Space(21); Mid$(xValorExt, 1, 58)
            Print #1, Space(21); Mid$(xValorExt, 59, 58)
            Print #1, Space(21); Mid$(xValorExt, 117, 58)
            
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            Print #1,
            
            Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
            Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
            Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
            Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            If de_informa.rsSel_Fatura.Fields("avulsa") <> "AVULSA" Then
            
                'BUSCA CTC A SER IMPRESSO
                If de_informa.rsSel_FaturaItens.State = 1 Then de_informa.rsSel_FaturaItens.Close
                de_informa.Sel_FaturaItens xFilialFatura
            
                xlin = 1
                xcol = 1
                
                Do Until de_informa.rsSel_FaturaItens.EOF
                
                    'Franklin Modifiquei abaixo ...
                    If de_informa.rsSel_FaturaItens.Fields("tipodoc") = "NFS" Then
                        Print #1, Spc(2); Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); " ";
                    Else
                        Print #1, Spc(1); Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); ' " ";
                    End If
                    'Fim
                    
                    If xcol = 3 Then 'terceira coluna
                        Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")
                        xlin = xlin + 1
                        xcol = 0
                        If xlin = 32 Then
                            
                            For xcont = 1 To 12
                                Print #1, ""
                            Next
                            
                            Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                                            mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                                            Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                            
                            Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
                            'Print #1, Space(4); "**" & "      "
                            Print #1, String(13 - Len(Trim$("********")), " "); "********"; "      ";
                            Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
                            Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
                            
                            Print #1, ""
                            Print #1, ""
                            
                            Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
                                
                            Print #1, ""
                                                    
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
                            Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
                            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
                            Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
                            Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
                            
                            'CGC + FORMATAÇÃO
                            If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
                                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
                            Else
                                xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
                            End If
                            
                            Print #1, Space(21); xcgc1;
                            Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
                            
                            Print #1,
                            Print #1,
                            Print #1, Space(21); "**"
                            'impreme o valor por extenso
                            
                            'xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
                            
                            'Print #1, Space(21); Mid$(xValorExt, 1, 58)
                            'Print #1, Space(21); Mid$(xValorExt, 59, 58)
                            'Print #1, Space(21); Mid$(xValorExt, 117, 58)
                            
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            Print #1,
                            
                            Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
                            Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
                            Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
                            Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
                            Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
                            
                            Print #1, ""
                            Print #1, ""
                            Print #1, ""
                                                
                            xlin = 1
                        
                        End If
                    Else
                        Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00"); " ";
                    End If
                    
                    xcol = xcol + 1
                    
                    de_informa.rsSel_FaturaItens.MoveNext
                                
                Loop
                
                For xcont = 1 To 32 - xlin
                    Print #1, ""
                Next
                
            Else
            
                For xcont = 1 To 31
                    Print #1, ""
                Next
            
            End If
            
        
            
            Print #1, ""
            Print #1, ""
            
            Print #1, " (-)ICMS:"; String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00");
            Print #1, "    (-)ABAT:"; String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("abatimento"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("abatimento"), "###,##0.00");
            Print #1, Space(25); String(14 - Len(Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Close #1
            
        End If
    End If
    
End Sub
Public Sub imprime_fatrel(xFilialFatura As String)
    Dim ximpr_inst As Printer, ximpr_cfg As String, xnumCTC As String, xlinha As String
    Dim xlin As Integer, xcol As Integer, xcont As Integer, xValorExt As String

    'busca impressora para este documento
    If Dir("C:\informa.cfg") <> "" Then
        
        Open "C:\informa.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "REL" Then
                If Trim$(Mid$(xlinha, 5, 2)) = "\\" Or Trim$(Mid$(xlinha, 5, 2)) = "**" Then
                    ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Else
                    ximpr_cfg = "LPT1:"
                End If
                Exit Do
            End If
        Loop
        
        If EOF(1) Then
            MsgBox "Não está Configurado a Impressora para Este Documento: CTC " & xFilialCtc
            Close #1
            Exit Sub
        End If
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: CTC " & xFilialCtc
        Exit Sub
        
    End If

    'BUSCA CTC A SER IMPRESSO
    If de_informa.rsSel_Fatura.State = 1 Then de_informa.rsSel_Fatura.Close
    de_informa.Sel_Fatura xFilialFatura
    
    If de_informa.rsSel_Fatura.RecordCount < 1 Then
        MsgBox "Fatura Inexistente !", vbInformation
        Exit Sub
    End If
    
    'seta impressora
    
'    If ximpr_cfg = "LPT1" Then
        Open ximpr_cfg For Output As #1
        DoEvents
'    Else
'        For Each ximpr_inst In Printers
'            If ximpr_inst.DeviceName = ximpr_cfg Then
'                Open ximpr_cfg For Output As #1
'                DoEvents
'                Exit For
'            End If
'        Next
'    End If

    'inicia a impressão
    Print #1, "RELATÓRIO DE FATURAMENTO"
    Print #1, ""
    Print #1, "FATURA: "; de_informa.rsSel_Fatura.Fields("filialfatura")
    Print #1, ""
    Print #1, ""
    
    Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_Fatura.Fields("emissao")), 2)); "   "; _
                    mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                    Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
    Print #1, String(13 - Len(Trim$(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"))), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"); "      ";
    Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
    Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
    
    Print #1, ""
    Print #1, ""

    Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
        
    Print #1, ""
        
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
    Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
    Print #1, Space(2); de_informa.rsSel_Fatura.Fields("cliente_uf")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("cliente_cidade")), " ");
    Print #1, Space(15); de_informa.rsSel_Fatura.Fields("cepcob")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_cgc");
    Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
    
    Print #1,
    
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    
    Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
    Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
    Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
    Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
    Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    If de_informa.rsSel_Fatura.Fields("avulsa") = "N" Then
    
        'BUSCA CTC A SER IMPRESSO
        If de_informa.rsSel_FaturaItens.State = 1 Then de_informa.rsSel_FaturaItens.Close
        de_informa.Sel_FaturaItens xFilialFatura
    
        xlin = 1
        xcol = 1
        
        Do Until de_informa.rsSel_FaturaItens.EOF
        
            Print #1, Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); " ";
            
            If xcol = 3 Then 'terceira coluna
                Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")
                xlin = xlin + 1
                xcol = 0
                If xlin = 32 Then
                    For xcont = 1 To 34
                        Print #1, ""
                    Next
                    xlin = 1
                End If
            Else
                Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00"); " ";
            End If
            
            xcol = xcol + 1
            
            de_informa.rsSel_FaturaItens.MoveNext
                        
        Loop
        
        For xcont = 1 To 32 - xlin
            Print #1, ""
        Next
        
    Else
    
        For xcont = 1 To 31
            Print #1, ""
        Next
    
    End If
    

    
    Print #1, ""
    Print #1, ""
    
    Print #1, " (-)ICMS:"; String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00");
    Print #1, "    (-)ABAT:"; String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("abatimento"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("abatimento"), "###,##0.00");
    Print #1, Space(25); String(14 - Len(Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Close #1
    
    
End Sub
Public Sub impr_prefat(xfilialprefat As String)
    Dim ximpr_inst As Printer, ximpr_cfg As String, xnumCTC As String, xlinha As String
    Dim xlin As Integer, xcol As Integer, xcont As Integer, xValorExt As String

    'busca impressora para este documento
    If Dir("C:\informa.cfg") <> "" Then
        
        Open "C:\informa.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "FAT" Then
                If Trim$(Mid$(xlinha, 5, 2)) = "\\" Then
                    ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Else
                    ximpr_cfg = "LPT1:"
                End If
                Exit Do
            End If
        Loop
        
        If EOF(1) Then
            MsgBox "Não está Configurado a Impressora para Este Documento: CTC " & xFilialCtc
            Close #1
            Exit Sub
        End If
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: CTC " & xFilialCtc
        Exit Sub
        
    End If

    'BUSCA CTC A SER IMPRESSO
    If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
    de_informa.Sel_PreFatura xfilialprefat
    
    If de_informa.rsSel_PreFatura.RecordCount < 1 Then
        MsgBox "Pré-Fatura Inexistente !", vbInformation
        Exit Sub
    End If
    
    'seta impressora
    
'    If ximpr_cfg = "LPT1" Then
        Open ximpr_cfg For Output As #1
        DoEvents
'    Else
'        For Each ximpr_inst In Printers
'            If ximpr_inst.DeviceName = ximpr_cfg Then
'                Open ximpr_cfg For Output As #1
'                DoEvents
'                Exit For
'            End If
'        Next
'    End If

    'inicia a impressão
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Print #1, Space(58); Trim$(zeros(Day(de_informa.rsSel_PreFatura.Fields("emissao")), 2)); "   "; _
                    mesnome(Month(de_informa.rsSel_PreFatura.Fields("emissao"))); String(12 - Len(mesnome(Month(de_informa.rsSel_Fatura.Fields("emissao")))), " "); _
                    Trim$(Str(Year(de_informa.rsSel_Fatura.Fields("emissao"))))
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Print #1, Space(13); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "    ";
    Print #1, String(13 - Len(Trim$(Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"))), " "); Format(de_informa.rsSel_Fatura.Fields("valorfatura"), "##,###,##0.00"); "      ";
    Print #1, Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6); "     ";
    Print #1, de_informa.rsSel_Fatura.Fields("vencimento")
    
    Print #1, ""
    Print #1, ""

    Print #1, Space(40); de_informa.rsSel_Fatura.Fields("banconome")
        
    Print #1, ""
        
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_nome")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cliente_end")
    Print #1, Space(21); Space(30); Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20); String(20 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_cidade"), 1, 20)), " ");
    Print #1, Space(5); de_informa.rsSel_Fatura.Fields("cliente_uf")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("cidadecob"); "-"; de_informa.rsSel_Fatura.Fields("ufcob")
    Print #1, Space(21); de_informa.rsSel_Fatura.Fields("endcob"); String(40 - Len(de_informa.rsSel_Fatura.Fields("endcob")), " ");
    Print #1, Space(10); de_informa.rsSel_Fatura.Fields("cepcob")
    
    'CGC + FORMATAÇÃO
    If Len(Trim$(Str(CDbl(de_informa.rsSel_Fatura.Fields("cliente_cgc"))))) > 11 Then
        xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@.@@@.@@@/@@@@-@@")
    Else
        xcgc1 = Format(de_informa.rsSel_Fatura.Fields("cliente_cgc"), "@@@.@@@.@@@-@@")
    End If
    
    Print #1, Space(21); xcgc1;
    Print #1, Space(11); de_informa.rsSel_Fatura.Fields("cliente_ie")
    
    Print #1,
    
    'impreme o valor por extenso
    
    xValorExt = Extenso(de_informa.rsSel_Fatura.Fields("valorfatura"), "REAIS", "REAL", 174)
    
    Print #1, Space(21); Mid$(xValorExt, 1, 58)
    Print #1, Space(21); Mid$(xValorExt, 59, 58)
    Print #1, Space(21); Mid$(xValorExt, 117, 58)
    
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    
    Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
    Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
    Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
    Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
    Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    If de_informa.rsSel_Fatura.Fields("avulsa") = "N" Then
    
        'BUSCA CTC A SER IMPRESSO
        If de_informa.rsSel_FaturaItens.State = 1 Then de_informa.rsSel_FaturaItens.Close
        de_informa.Sel_FaturaItens xFilialFatura
    
        xlin = 1
        xcol = 1
        
        Do Until de_informa.rsSel_FaturaItens.EOF
        
            Print #1, Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 1, 2); "-"; Mid$(de_informa.rsSel_FaturaItens.Fields("filialctc"), 3, 8); " ";
            
            If xcol = 3 Then 'terceira coluna
                Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")
                xlin = xlin + 1
                xcol = 0
                If xlin = 32 Then
                    For xcont = 1 To 42
                        Print #1, ""
                    Next
                    Print #1, "  "; Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 35); String(35 - Len(Mid$(de_informa.rsSel_Fatura.Fields("cliente_nome"), 1, 30)), " ");
                    Print #1, Space(6); de_informa.rsSel_Fatura.Fields("cliente_cidade")
                    Print #1, Space(6); Mid$(de_informa.rsSel_Fatura.Fields("filialfatura"), 3, 6);
                    Print #1, Space(12); de_informa.rsSel_Fatura.Fields("emissao");
                    Print #1, Space(5); de_informa.rsSel_Fatura.Fields("vencimento")
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    xlin = 1
                End If
            Else
                Print #1, String(14 - Len(Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00")), " "); Format(de_informa.rsSel_FaturaItens.Fields("fretebruto"), "##,###,##0.00"); " ";
            End If
            
            xcol = xcol + 1
            
            de_informa.rsSel_FaturaItens.MoveNext
                        
        Loop
        
        For xcont = 1 To 32 - xlin
            Print #1, ""
        Next
        
    Else
    
        For xcont = 1 To 31
            Print #1, ""
        Next
    
    End If
    

    
    Print #1, ""
    Print #1, ""
    
    Print #1, " (-)ICMS:"; String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("descicms"), "###,##0.00");
    Print #1, "    (-)ABAT:"; String(10 - Len(Format(de_informa.rsSel_Fatura.Fields("abatimento"), "###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("abatimento"), "###,##0.00");
    Print #1, Space(25); String(14 - Len(Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")), " "); Format(de_informa.rsSel_Fatura.Fields("valorbrutoicms"), "##,###,##0.00")
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Close #1
    
    

End Sub
