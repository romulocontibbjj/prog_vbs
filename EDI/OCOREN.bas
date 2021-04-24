Attribute VB_Name = "OCOREN"
Public Function OCORENARQ(txtcgc As String, xarquivo As String)
Dim xlinha As String, xRemet_nome As String, xdata As String, xhora As String, xid_intercam As String
    Dim xNumNF As String, xobs_ocorr As String, xdataoco As String, xhoraoco As String, xserie As String
    
    
       
    If Mid$(Trim$(txtcgc), 1, 8) = "04229761" Then 'VIDEOLAR
    
        If de_informa.rsSel_EDI_Ocorr.State = 1 Then de_informa.rsSel_EDI_Ocorr.Close
        de_informa.Sel_EDI_Ocorr Trim$(txtcgc), "S"
                             
    Else
    
        If de_informa.rsSel_EDI_Ocorr.State = 1 Then de_informa.rsSel_EDI_Ocorr.Close
        de_informa.Sel_EDI_Ocorr Trim$(txtcgc), "%"
    
    End If
    
    If de_informa.rsSel_EDI_Ocorr.RecordCount > 0 Then
    
        'definição do diretório de gravação
        
          If txtcgc = "04229761" Then  'VIDEOLAR
            
            'PEGA O NUMERO MAX + 1 DE REGISTRO
            de_informa.rsSel_maxarq.Open
              
            
                xnumarq = Trim(de_informa.rsSel_maxarq.Fields("MAX"))
                de_informa.rsSel_maxarq.Close
                xnomearquivo = "LFTOCO" & String(4 - Len(xnumarq), "0") & xnumarq & ".TXT"
                
            'CRIO O LOCAL NO QUAL SERÁ SALVO O ARQ.
            
            If Mid$(Trim$(txtcgc), 1, 8) = "04229761" Then 'VIDEOLAR
            
                xarquivo = xarquivo & "\" & xnomearquivo
            
            Else
               
               xarquivo = xarquivo & xnomearquivo
                        
            End If
            
            
            'ALTERA TB_MEM
             de_informa.up_tbmem xnomearquivo, xnumarq, de_informa.rsSel_EDI_Ocorr.RecordCount
        
            frm_verifica.xvideolar = xarquivo
                    
         End If

        
        Open xarquivo For Output As #1
        
        'tratamentos de dados para o arquivo (cabecários)
        
        'nome remetente / embarcador
        xRemet_nome = Trim$(de_informa.rsSel_EDI_Ocorr.Fields("remet_nome"))
        
        If Len(xRemet_nome) > 35 Then
            xRemet_nome = Mid$(xRemet_nome, 1, 35)
        ElseIf Len(xRemet_nome) < 35 Then
            xRemet_nome = xRemet_nome + Space(35 - Len(xRemet_nome))
        End If
        
        'data
        xdata = ""
        If Len(Trim$(Str(Day(datahora("data"))))) = 1 Then
            xdata = xdata & "0" & Trim$(Str(Day(datahora("data"))))
        Else
            xdata = xdata & Trim$(Str(Day(datahora("data"))))
        End If
        If Len(Trim$(Str(Month(datahora("data"))))) = 1 Then
            xdata = xdata & "0" & Trim$(Str(Month(datahora("data"))))
        Else
            xdata = xdata & Trim$(Str(Month(datahora("data"))))
        End If
        xdata = xdata & Trim$(Str(Year(datahora("data"))))
        
        'hora
        xhora = Mid(Trim$(Str(Time())), 1, 2) & Mid(Trim$(Str(Time())), 4, 2)
        
        'identif. de intercambio
        xid_intercam = "OCO" & Mid(xdata, 1, 4) & Mid(xhora, 1, 4) & "0"
        
        'REGISTRO 000
        
        xlinha = "000INTEC TRANSPORTES                  " & xRemet_nome & Mid(xdata, 1, 4) & Mid$(xdata, 7, 2) & xhora & xid_intercam & Space(25)
        Print #1, UCase(xlinha)
        
        'REGISTRO 340
        
        xlinha = "340OCORR" & Mid$(xid_intercam, 4, 9) & Space(103)
        Print #1, UCase(xlinha)
        
        'REGISTRO 341
        
        xlinha = "34152134798000320INTEC INTEGRACAO NAC TRANSP ENC. CARGAS" & Space(64)
        Print #1, xlinha
    
        'inicio do laço na recordset
        
        Do Until de_informa.rsSel_EDI_Ocorr.EOF
            
            If Mid$(txtcgc, 1, 8) = "50929710" And Not (Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "50929710" Or _
                                    Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "04019475" Or _
                                    Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "52134798" Or _
                                    Mid$(de_informa.rsSel_EDI_Ocorr.Fields("respons_cgc"), 1, 8) = "02426290") Then
                de_informa.rsSel_EDI_Ocorr.MoveNext
            Else

                'tratamento dos dados do detalhe (ocorrência)
            
                'número da NF
                xNumNF = String(8 - Len(Trim$(de_informa.rsSel_EDI_Ocorr.Fields("numnf"))), "0") & Trim$(de_informa.rsSel_EDI_Ocorr.Fields("numnf"))
        
                'observação de ocorrência
                If Not IsNull(de_informa.rsSel_EDI_Ocorr.Fields("obs_ocorr")) Then
                    xobs_ocorr = Trim$(de_informa.rsSel_EDI_Ocorr.Fields("obs_ocorr"))
                Else
                    xobs_ocorr = Space(70)
                End If
                If Len(xobs_ocorr) > 70 Then
                    xobs_ocorr = Mid$(xobs_ocorr, 1, 70)
                ElseIf Len(xobs_ocorr) < 70 Then
                    xobs_ocorr = xobs_ocorr + Space(70 - Len(xobs_ocorr))
                End If
            
                'data ocorrência
                xdataoco = ""
                If Len(Trim$(Str(Day(de_informa.rsSel_EDI_Ocorr.Fields("data"))))) = 1 Then
                    xdataoco = xdataoco & "0" & Trim$(Str(Day(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                Else
                    xdataoco = xdataoco & Trim$(Str(Day(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                End If
                If Len(Trim$(Str(Month(de_informa.rsSel_EDI_Ocorr.Fields("data"))))) = 1 Then
                    xdataoco = xdataoco & "0" & Trim$(Str(Month(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                Else
                    xdataoco = xdataoco & Trim$(Str(Month(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
                End If
                xdataoco = xdataoco & Trim$(Str(Year(de_informa.rsSel_EDI_Ocorr.Fields("data"))))
            
                'hora ocorrência
                xhoraoco = Mid$(de_informa.rsSel_EDI_Ocorr.Fields("hora"), 1, 2) & Mid$(de_informa.rsSel_EDI_Ocorr.Fields("hora"), 4, 2)
                If xhoraoco = "" Then
                    xhoraoco = "0000"
                End If
            
                'SERIE DA NF
                If Mid$(txtcgc, 1, 8) = "50929710" Then     'medley
                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
                        xserie = "   "
                    Else
                        xserie = Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))) & _
                                 String(3 - (Len(Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))))), " ")
                    End If
                ElseIf Mid$(txtcgc, 1, 8) = "04490850" Then 'gillette
                    xserie = "001"
                ElseIf Mid$(txtcgc, 1, 8) = "14372981" Then 'bayer
                    xserie = "001"
                ElseIf Mid$(txtcgc, 1, 8) = "61188488" Then 'bayer
                    xserie = "004"
                ElseIf Mid$(txtcgc, 1, 8) = "60412327" Then 'ALCON
                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
                        xserie = "01 "
                    Else
                        xserie = zeros(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")), 2) & " "
                    End If
                ElseIf Mid$(txtcgc, 1, 8) = "33247743" Then 'GLAXO
'                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
'                        xserie = "01 "
'                    Else
'                       xserie = zeros(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")), 2) & " "
'                    End If
                    xserie = "01 "
                Else
                    If Not IsNumeric(de_informa.rsSel_EDI_Ocorr.Fields("serie")) Then
                        xserie = "   "
                    Else
                        xserie = Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))) & _
                                 String(3 - (Len(Trim$(Str(CDbl(de_informa.rsSel_EDI_Ocorr.Fields("serie")))))), " ")
                    End If
                End If
                
                'REGISTRO 342

                xlinha = "342" & de_informa.rsSel_EDI_Ocorr.Fields("remet_cgc") & xserie & xNumNF & de_informa.rsSel_EDI_Ocorr.Fields("cod_ocorr") & _
                        xdataoco & xhoraoco & "00" & xobs_ocorr & Space(6)
                Print #1, xlinha
            
                'ATUALIZA EDI GERADO = S
            
                de_informa.Alt_EDI_Ocorr de_informa.rsSel_EDI_Ocorr.Fields("codigo")
            
                de_informa.rsSel_EDI_Ocorr.MoveNext
                
            End If
            
        Loop
        Close #1
    
        OCORENARQ = 1
    
    
    Else
        OCORENARQ = 0
    End If





End Function
