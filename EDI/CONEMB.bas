Attribute VB_Name = "CONEMB"
Public Function CONEMB1(cliente As String, txtCGCRem As String, xnomearquivo As String, periodo As Integer, xentr As Integer, xcanc As Integer)
Dim xRegID As String
Dim xAgora As Date
Dim xNomeFile As String
Dim Filial As String
Dim xserie As String
Dim CTC As String
Dim xdata As String
Dim xCondicaoFrete As String
Dim xpeso As String
Dim FreteTotal As String
Dim BaseCalc As String
Dim Aliquota As String
Dim ICMS As String
Dim FreteVolume As String
Dim FreteValor As String
Dim SecCat As String
Dim ITR As String
Dim Despacho As String
Dim Pedagio As String
Dim AdEme As String
Dim Subst As String
Dim CFOP As String
Dim CgcEmissor As String
Dim CgcEmbarc As String
Dim xAcao As String
Dim xTipoCon As String
Dim xFiller As String

Dim xRemID As String
Dim xDesID As String

Dim xDia As String
Dim xmes As String
Dim xano As String

Dim xH As String
Dim xM As String
Dim xhora As String
Dim xIntID As String
Dim xlinha As String
Dim xDocID As String
Dim xcgc As String
Dim xRazaoSocial As String

Dim xRecSize As Integer

Dim TotFrete As Long
Dim TotCTC As Long

Dim xTotFrete As String
Dim xTotCtc As String

Dim xAux As String

If xentr = 1 Then
    xmotivodoc = "ENT"
    xRemetCgc = Mid$(txtCGCRem, 1, 8) + "%"
Else
    xmotivodoc = "%"
    xRemetCgc = "%"
End If

If de_informa.rsCONEMBSel.State = 1 Then de_informa.rsCONEMBSel.Close
de_informa.CONEMBSel CDate(Date - periodo), CDate(Date), Trim(UCase(txtCGCRem)) & "%", _
                     "%", xRemetCgc, xmotivodoc
                     
If de_informa.rsCONEMBSel.RecordCount > 0 Then
XNARQ = Int((Hour(Time) + Day(Date) + Year(Date) + Minute(Time)) / Second(Time))

xAgora = datahora("DATAHORA")
xNomeFile = Mid$(Trim(UCase(txtCGCRem)), 1, 8) & "_" & zeros(Day(xAgora), 2) & zeros(Month(xAgora), 2) & "_" & zeros(Hour(xAgora), 2) & zeros(Minute(xAgora), 2) & ".txt"

    Open xnomearquivo For Output As #1

    
xRecSize = 680

xRegID = "000"

If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
de_informa.Sel_CadFilial de_informa.rsCONEMBSel.Fields("filial")

xRemID = Mid$(de_informa.rsSel_CadFilial.Fields("empresa"), 1, 35)
xDesID = UCase(cliente)

xDia = String(2 - Len(Trim(Str(Day(Date)))), "0") & Trim(Str(Day(Date)))
xmes = String(2 - Len(Trim(Str(Month(Date)))), "0") & Trim(Str(Month(Date)))
xano = String(2 - Len(Mid(Trim(Str(Year(Date))), 3, 2)), "0") & Mid(Trim(Str(Year(Date))), 3, 2)
xdata = xDia & xmes & xano

xH = String(2 - Len(Trim(Str(Hour(Time)))), "0") & Trim(Str(Hour(Time)))
xM = String(2 - Len(Trim(Str(Minute(Time)))), "0") & Trim(Str(Minute(Time)))
xhora = xH & xM

xIntID = "CON" & xDia & xmes & xhora & "0"

xRegID = String(3 - Len(Trim(xRegID)), " ") & Trim(xRegID)
xRemID = Trim(xRemID) & String(35 - Len(Trim(xRemID)), " ")
xDesID = Trim(xDesID) & String(35 - Len(Trim(xDesID)), " ")
xdata = Trim(xdata) & String(6 - Len(Trim(xdata)), " ")
xhora = Trim(xhora) & String(4 - Len(Trim(xhora)), " ")
xIntID = Trim(xIntID) & String(12 - Len(Trim(xIntID)), " ")
xFiller = String(585, " ")
xlinha = xRegID & xRemID & xDesID & xdata & xhora & xIntID & xFiller
    If Len(xlinha) <> xRecSize Then
    MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
    Close #1
    Else
    Print #1, xlinha
    End If
    
xRegID = "320"
xDocID = "CONHE" & xDia & xmes & xhora & "1"
xFiller = String(663, " ")
xlinha = xRegID & xDocID & xFiller
    If Len(xlinha) <> xRecSize Then
    MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
    Close #1
    Else
    Print #1, xlinha
    End If
    
xRegID = "321"
xcgc = de_informa.rsSel_CadFilial.Fields("cgc")
xRazaoSocial = de_informa.rsSel_CadFilial.Fields("empresa")

xRegID = Trim(xRegID) & String(3 - Len(Trim(xRegID)), " ")
xcgc = Trim(xcgc) & String(14 - Len(Trim(xcgc)), " ")
xRazaoSocial = Trim(xRazaoSocial) & String(40 - Len(Trim(xRazaoSocial)), " ")
xFiller = String(623, " ")
xlinha = xRegID & xcgc & xRazaoSocial & xFiller
    
    If Len(xlinha) <> xRecSize Then
    MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
    Close #1
    Else
    Print #1, xlinha
    End If
    
TotFrete = 0
TotCTC = 0

    Do Until de_informa.rsCONEMBSel.EOF
    
        If xcanc = 0 And de_informa.rsCONEMBSel.Fields("tem_ocorr") = "C" Then
            xprocessaregistro = "N"
        ElseIf xcanc = 1 And de_informa.rsCONEMBSel.Fields("tem_ocorr") <> "C" Then
            xprocessaregistro = "N"
        Else
            xprocessaregistro = "S"
        End If
        
        If xprocessaregistro = "S" Then
    
            TotFrete = TotFrete + de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
            TotCTC = TotCTC + 1
            
            xRegID = "322"
            
            If IsNull(de_informa.rsCONEMBSel.Fields("codfilial")) Or de_informa.rsCONEMBSel.Fields("codfilial") = "" Then
                Filial = de_informa.rsCONEMBSel.Fields("filial")
            Else
                Filial = de_informa.rsCONEMBSel.Fields("codfilial")
            End If
            
            If IsNull(de_informa.rsCONEMBSel.Fields("seriectc")) Or de_informa.rsCONEMBSel.Fields("seriectc") = "" Then
                xserie = "U"
            Else
                xserie = de_informa.rsCONEMBSel.Fields("seriectc")
            End If
            
            CTC = de_informa.rsCONEMBSel.Fields("CTC")
            xDia = String(2 - Len(Trim(Str(Day(de_informa.rsCONEMBSel.Fields("DATA"))))), "0") & Trim(Str(Day(de_informa.rsCONEMBSel.Fields("DATA"))))
            xmes = String(2 - Len(Trim(Str(Month(de_informa.rsCONEMBSel.Fields("DATA"))))), "0") & Trim(Str(Month(de_informa.rsCONEMBSel.Fields("DATA"))))
            xano = String(4 - Len(Trim(Str(Year(de_informa.rsCONEMBSel.Fields("DATA"))))), "0") & Trim(Str(Year(de_informa.rsCONEMBSel.Fields("DATA"))))
            xdata = xDia & xmes & xano
            xCondicaoFrete = de_informa.rsCONEMBSel.Fields("FPAG")
                If xCondicaoFrete = "1-CIF" Then xCondicaoFrete = "C"
                If xCondicaoFrete = "2-FOB" Then xCondicaoFrete = "F"
                If xCondicaoFrete = "A PAGAR" Then xCondicaoFrete = "F"
                If xCondicaoFrete = "AGO   G" Then xCondicaoFrete = "C"
                If xCondicaoFrete = "PAGO" Then xCondicaoFrete = "C"
                
            xpeso = de_informa.rsCONEMBSel.Fields("PESO")
            
            If Mid$(de_informa.rsCONEMBSel.Fields("respons_cgc"), 1, 8) = "04490850" Then
                If de_informa.rsCONEMBSel.Fields("subtrib") = "S" Then
                    FreteTotal = de_informa.rsCONEMBSel.Fields("FRETETOTAL")
                Else
                    FreteTotal = de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
                End If
            Else
                FreteTotal = de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
            End If
            
            BaseCalc = de_informa.rsCONEMBSel.Fields("FRETETOTALBRUTO")
            Aliquota = de_informa.rsCONEMBSel.Fields("ALIQUOTA") * 100
            ICMS = FreteTotal * (Aliquota / 100)
            FreteVolume = 0
            FreteValor = de_informa.rsCONEMBSel.Fields("FRETEVALORBR")
            SecCat = de_informa.rsCONEMBSel.Fields("TXCOLETABR") + de_informa.rsCONEMBSel.Fields("TXENTREGAredbr")
            ITR = "0"
            Despacho = "0"
            Pedagio = de_informa.rsCONEMBSel.Fields("PEDAGIOBR")
            AdEme = "0"
            Subst = de_informa.rsCONEMBSel.Fields("SUBTRIB")
                If Subst = "S" Then Subst = "1"
                If Subst = "N" Then Subst = "2"
            CFOP = Mid(de_informa.rsCONEMBSel.Fields("CFOP"), 1, 3)
            CgcEmissor = de_informa.rsSel_CadFilial.Fields("cgc")
            CgcEmbarc = de_informa.rsCONEMBSel.Fields("REMET_CGC")
            
            If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
            de_informa.Sel_NFsdoCTC transctc(de_informa.rsCONEMBSel.Fields("filial"), CTC)
                
            Dim NFs(1 To 40, 1 To 2) As String
                For k = 1 To 40
                    If de_informa.rsSel_NFsdoCTC.EOF = False Then
                        NFs(k, 1) = de_informa.rsSel_NFsdoCTC.Fields("SERIE")
                        NFs(k, 2) = de_informa.rsSel_NFsdoCTC.Fields("NUMNF")
                        de_informa.rsSel_NFsdoCTC.MoveNext
                    Else
                        NFs(k, 1) = ""
                        NFs(k, 2) = ""
                    End If
                    
                    If Val(NFs(k, 1)) > 0 Then
                        NFs(k, 1) = Trim(Str(Val(NFs(k, 1)))) & String(3 - Len(Trim(Str(Val(NFs(k, 1))))), " ")
                    Else
                        If Val(NFs(k, 1)) = 0 And Val(NFs(k, 2)) > 0 Then
                            NFs(k, 1) = "U  "
                        Else
                            NFs(k, 1) = "   "
                        End If
                    End If
                    
                    NFs(k, 2) = String(8 - Len(Trim(Str(Val(NFs(k, 2))))), "0") & Trim(Str(Val(NFs(k, 2))))
                    
                Next
                
            Dim StringNFs As String
            
            StringNFs = ""
            
                For k = 1 To 40
                    StringNFs = StringNFs & NFs(k, 1) & NFs(k, 2)
                Next
            
            'VIDEOLAR CONTI: 22/02 13:40
            
'            xAux = "1  "
'            xAux = xAux & Mid(StringNFs, 4, Len(StringNFs) - 3)
'            StringNFs = xAux
            
            If de_informa.rsCONEMBSel.Fields("tem_ocorr") = "C" Then
                xAcao = "E"
            Else
                xAcao = "I"
            End If
            
            xTipoCon = "N"
                
            xpeso = Trim(Str(Int(Val(xpeso * 100))))
            FreteTotal = Trim(Str(Int(Val(FreteTotal * 100))))
            BaseCalc = Trim(Str(Int(Val(BaseCalc * 100))))
            Aliquota = Trim(Str(Int(Val(Aliquota * 100))))
            ICMS = Trim(Str(Int(Val(xicms * 100))))
            FreteVolume = Trim(Str(Int(Val(FreteVolume * 100))))
            FreteValor = Trim(Str(Int(Val(FreteValor * 100))))
            SecCat = Trim(Str(Int(Val(SecCat * 100))))
            ITR = Trim(Str(Int(Val(ITR * 100))))
            Despacho = Trim(Str(Int(Val(Despacho * 100))))
            Pedagio = Trim(Str(Int(Val(pedadio * 100))))
            AdEme = Trim(Str(Int(Val(AdEme * 100))))
            
            xRegID = Trim(xRegID)
            Filial = Trim(Filial)
            xserie = Trim(xserie)
            CTC = Trim(CTC)
            xdata = Trim(xdata)
            xCondicaoFrete = Trim(xCondicaoFrete)
            xpeso = Trim(xpeso)
            FreteTotal = Trim(FreteTotal)
            BaseCalc = Trim(BaseCalc)
            Aliquota = Trim(Aliquota)
            ICMS = Trim(ICMS)
            FreteVolume = Trim(FreteVolume)
            FreteValor = Trim(FreteValor)
            SecCat = Trim(SecCat)
            ITR = Trim(ITR)
            Despacho = Trim(Despacho)
            Pedagio = Trim(Pedagio)
            AdEme = Trim(AdEme)
            Subst = Trim(Subst)
            CFOP = Trim(CFOP)
            CgcEmissor = Trim(CgcEmissor)
            CgcEmbarc = Trim(CgcEmbarc)
            xAcao = Trim(xAcao)
            xTipoCon = Trim(xTipoCon)
            
            xRegID = xRegID & String(3 - Len(xRegID), " ")
            Filial = Filial & String(10 - Len(Filial), " ")
        
            
            xserie = xserie & String(5 - Len(xserie), " ")
         
            
            CTC = CTC & String(12 - Len(CTC), " ")
            xdata = xdata & String(8 - Len(xdata), " ")
            xCondicaoFrete = xCondicaoFrete & String(1 - Len(xCondicaoFrete), " ")
            xpeso = String(7 - Len(xpeso), "0") & xpeso
            FreteTotal = String(15 - Len(FreteTotal), "0") & FreteTotal
            BaseCalc = String(15 - Len(BaseCalc), "0") & BaseCalc
            Aliquota = String(4 - Len(Aliquota), "0") & Aliquota
            ICMS = String(15 - Len(ICMS), "0") & ICMS
            FreteVolume = String(15 - Len(FreteVolume), "0") & FreteVolume
            FreteValor = String(15 - Len(FreteValor), "0") & FreteValor
            SecCat = String(15 - Len(SecCat), "0") & SecCat
            ITR = String(15 - Len(ITR), "0") & ITR
            Despacho = String(15 - Len(Despacho), "0") & Despacho
            Pedagio = String(15 - Len(Pedagio), "0") & Pedagio
            AdEme = String(15 - Len(AdEme), "0") & AdEme
            Subst = Subst & String(1 - Len(Subst), " ")
            CFOP = CFOP & String(3 - Len(CFOP), " ")
            
            If txtCGCRem = "04229761" Then
                
                CgcEmissor = "52134798000168"
                
            Else
              
                CgcEmissor = CgcEmissor & String(14 - Len(CgcEmissor), " ")
                
            End If
            
            CgcEmbarc = CgcEmbarc & String(14 - Len(CgcEmbarc), " ")
            xAcao = xAcao & String(1 - Len(xAcao), " ")
            xTipoCon = xTipoCon & String(1 - Len(xTipoCon), " ")
            xFiller = String(6, " ")
        
            xlinha = xRegID & Filial & xserie & CTC & xdata & xCondicaoFrete & xpeso & FreteTotal & BaseCalc & Aliquota & ICMS & FreteVolume & FreteValor & SecCat & ITR & Despacho & Pedagio & AdEme & Subst & CFOP & CgcEmissor & CgcEmbarc & StringNFs & xAcao & xTipoCon & xFiller
                
                If Len(xlinha) <> xRecSize Then
                MsgBox "Existe um erro de compilaçao no registro " & xRegID & ". Por favor, verifique seu código.", vbCritical
                Close #1
                Else
                Print #1, xlinha
                End If
                
            End If
            
            de_informa.rsCONEMBSel.MoveNext
            DoEvents
            
    Loop


xRegID = "323"
xTotFrete = Trim(Str(Int(TotFrete * 100)))
xTotCtc = Trim(Str(TotCTC))

If CDbl(xTotCtc) > 9999 Then
    xTotCtc = "0000"
Else
    xTotCtc = String(4 - Len(xTotCtc), "0") & xTotCtc
End If
xTotFrete = String(15 - Len(xTotFrete), "0") & xTotFrete
xFiller = String(658, " ")

xlinha = xRegID & xTotCtc & xTotFrete & xFiller

Print #1, xlinha
    
Close #1

de_informa.rsCONEMBSel.MoveFirst

Do Until de_informa.rsCONEMBSel.EOF
    'atualiza at-edi-cif indicando que já foi enviado o EDI de CTC para o cliente
    
    If xcanc = 1 And de_informa.rsCONEMBSel.Fields("tem_ocorr") = "C" Then
        xprocessaregistro = "N"
    ElseIf xcanc = 1 And de_informa.rsCONEMBSel.Fields("tem_ocorr") <> "C" Then
        xprocessaregistro = "N"
    Else
        xprocessaregistro = "S"
    End If
    
    If xprocessaregistro = "S" Then
        de_informa.Alt_EDIConembSIM de_informa.rsCONEMBSel.Fields("filialctc")
    End If
    
    DoEvents
    
    de_informa.rsCONEMBSel.MoveNext
    
Loop

CONEMB1 = 1
Else
CONEMB1 = 0
End If


End Function
