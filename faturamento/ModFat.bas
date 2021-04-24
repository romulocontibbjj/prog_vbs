Attribute VB_Name = "ModFat"
Option Explicit
Public contador As Integer
Public xusuario As String
Public xdireitos As String
Public xamarelo1 As Long
Public xamarelo2 As Long
Public xbranco As Long
Public xstrcon As String
Public xstrcon2 As String
Public xCnx As String
Public xBco As String

Public Function transctc(xfilial As String, xctc As String) As String
    Dim X As Integer
    X = Len(Trim(xctc))
    If X = 1 Then
        xctc = "0000000" & Trim(xctc)
    ElseIf X = 2 Then
        xctc = "000000" & Trim(xctc)
    ElseIf X = 3 Then
        xctc = "00000" & Trim(xctc)
    ElseIf X = 4 Then
        xctc = "0000" & Trim(xctc)
    ElseIf X = 5 Then
        xctc = "000" & Trim(xctc)
    ElseIf X = 6 Then
        xctc = "00" & Trim(xctc)
    ElseIf X = 7 Then
        xctc = "0" & Trim(xctc)
    End If
    transctc = xfilial & Trim(xctc)
End Function
Public Function TransFatur(xfilial As String, xFatura As String) As String
    Dim X As Integer
    X = Len(Trim(xFatura))
    If X = 1 Then
        xFatura = "00000" & Trim(xFatura)
    ElseIf X = 2 Then
        xFatura = "0000" & Trim(xFatura)
    ElseIf X = 3 Then
        xFatura = "000" & Trim(xFatura)
    ElseIf X = 4 Then
        xFatura = "00" & Trim(xFatura)
    ElseIf X = 5 Then
        xFatura = "0" & Trim(xFatura)
    End If
    TransFatur = xfilial & Trim(xFatura)
End Function
Public Function diasemana(xdata As Date)
    If Weekday(xdata) = 1 Then
        diasemana = "Domingo"
    ElseIf Weekday(xdata) = 2 Then
        diasemana = "Segunda-Feira"
    ElseIf Weekday(xdata) = 3 Then
        diasemana = "Terca-Feira"
    ElseIf Weekday(xdata) = 4 Then
        diasemana = "Quarta-Feira"
    ElseIf Weekday(xdata) = 5 Then
        diasemana = "Quinta-Feira"
    ElseIf Weekday(xdata) = 6 Then
        diasemana = "Sexta-Feira"
    ElseIf Weekday(xdata) = 7 Then
        diasemana = "Sabado"
    End If
End Function
Public Sub tempo(xtempo As Integer)
Dim xtimer As Long
xtimer = Timer()
Do While True
    If Timer() >= xtimer + xtempo Then
        Exit Sub
    End If
Loop
End Sub
Public Function century(xdata As String) As String
    If xdata <> "__/__/____" Then
        If Mid$(xdata, 9, 2) = "__" Then
            If Val(Mid$(xdata, 7, 2)) >= 90 And Val(Mid$(xdata, 7, 2)) <= 99 Then
                century = Mid$(xdata, 1, 6) & "19" & Mid$(xdata, 7, 2)
            ElseIf Val(Mid$(xdata, 7, 2)) >= 0 And Val(Mid$(xdata, 7, 2)) < 90 Then
                century = Mid$(xdata, 1, 6) & "20" & Mid$(xdata, 7, 2)
            End If
        Else
            century = xdata
        End If
    Else
            century = xdata
    End If
End Function
Public Function zeros(xnumero As Long, xqtde As Integer) As String
    zeros = String(xqtde - Len(Trim$(Str(xnumero))), "0") & Trim$(Str(xnumero))
End Function
Public Sub limpatela(xtela As Form)
    Dim xmask As String
    Dim xcontrol As Control
    For Each xcontrol In xtela.Controls
        If TypeOf xcontrol Is TextBox Then
            xcontrol.Text = ""
        ElseIf TypeOf xcontrol Is Label Then
            If xcontrol.BorderStyle = 1 Then
                xcontrol.Caption = ""
            End If
        ElseIf TypeOf xcontrol Is MaskEdBox Then
            xmask = xcontrol.Mask
            xcontrol.Mask = ""
            xcontrol.Text = ""
            xcontrol.Mask = xmask
        End If
    Next
End Sub
Public Function SoNumeros(xnumero As String) As String
Dim xposicao As Integer, xstring As String
xposicao = 1
xstring = ""
Do While Len(xnumero) >= xposicao
    If IsNumeric(Mid$(xnumero, xposicao, 1)) Then
        xstring = xstring & Mid$(xnumero, xposicao, 1)
    End If
    xposicao = xposicao + 1
Loop
SoNumeros = xstring
End Function
Public Function MesAno(xmes As Long, xano As Long) As String
Dim xmesano As String
If xmes = 1 Then
    xmesano = "Jan/" & Trim$(Str(xano))
ElseIf xmes = 2 Then
    xmesano = "Fev/" & Trim$(Str(xano))
ElseIf xmes = 3 Then
    xmesano = "Mar/" & Trim$(Str(xano))
ElseIf xmes = 4 Then
    xmesano = "Abr/" & Trim$(Str(xano))
ElseIf xmes = 5 Then
    xmesano = "Mai/" & Trim$(Str(xano))
ElseIf xmes = 6 Then
    xmesano = "Jun/" & Trim$(Str(xano))
ElseIf xmes = 7 Then
    xmesano = "Jul/" & Trim$(Str(xano))
ElseIf xmes = 8 Then
    xmesano = "Ago/" & Trim$(Str(xano))
ElseIf xmes = 9 Then
    xmesano = "Set/" & Trim$(Str(xano))
ElseIf xmes = 10 Then
    xmesano = "Out/" & Trim$(Str(xano))
ElseIf xmes = 11 Then
    xmesano = "Nov/" & Trim$(Str(xano))
ElseIf xmes = 12 Then
    xmesano = "Dez/" & Trim$(Str(xano))
End If
MesAno = xmesano
End Function
Public Function datahora(xparametro As String) As Variant
    Dim xretorno As Variant
    If UCase(xparametro) = "DATA" Then
        If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
        de_informa.Sel_DataServidor
        xretorno = CDate(Trim$(Str(Year(de_informa.rsSel_DataServidor.Fields("agora")))) & "/" & _
                         Trim$(zeros2(Str(Month(de_informa.rsSel_DataServidor.Fields("agora"))), 2)) & "/" & _
                         Trim$(zeros2(Str(Day(de_informa.rsSel_DataServidor.Fields("agora"))), 2)))
        datahora = xretorno
    ElseIf UCase(xparametro) = "HORA" Then
        If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
        de_informa.Sel_DataServidor
        xretorno = zeros(Hour(de_informa.rsSel_DataServidor.Fields("agora")), 2) & ":" & _
                   zeros(Minute(de_informa.rsSel_DataServidor.Fields("agora")), 2) & ":" & _
                   zeros(Second(de_informa.rsSel_DataServidor.Fields("agora")), 2)
        datahora = xretorno
    ElseIf UCase(xparametro) = "DATAHORA" Then
        If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
        de_informa.Sel_DataServidor
        xretorno = CDate(Trim$(Str(Year(de_informa.rsSel_DataServidor.Fields("agora")))) & "/" & _
                         Trim$(zeros2(Str(Month(de_informa.rsSel_DataServidor.Fields("agora"))), 2)) & "/" & _
                         Trim$(zeros2(Str(Day(de_informa.rsSel_DataServidor.Fields("agora"))), 2)))
        xretorno = xretorno & " " & zeros(Hour(de_informa.rsSel_DataServidor.Fields("agora")), 2) & ":" & _
                         zeros(Minute(de_informa.rsSel_DataServidor.Fields("agora")), 2) & ":" & _
                         zeros(Second(de_informa.rsSel_DataServidor.Fields("agora")), 2)
        datahora = xretorno
    Else
        MsgBox "ERRO ! Parâmetro Inválido na Função DATAHORA() !", vbCritical, "ERRO de SISTEMA"
        datahora = ""
    End If
End Function
Public Sub combomesano(xcomboname As ComboBox)
    Dim xyear As Long, xmonth As Long, xyearcont As Long, xmonthcont As Long
    
    xyear = 2002 'ano inicial
    xmonth = 1   'mes inicial
    
    xcomboname.Clear
    xyearcont = Year(datahora("DATA"))
    xmonthcont = Month(datahora("DATA"))
    Do While True
        xcomboname.AddItem MesAno(xmonthcont, xyearcont)
        xcomboname.ItemData(xcomboname.NewIndex) = zeros(xyearcont, 4) & zeros(xmonthcont, 2)
        If zeros(xyearcont, 4) & zeros(xmonthcont, 2) <> zeros(xyear, 4) & zeros(xmonth, 2) Then
            If xmonthcont = 1 Then
                xmonthcont = 12
                xyearcont = xyearcont - 1
            Else
                xmonthcont = xmonthcont - 1
            End If
        Else
            Exit Do
        End If
    Loop
End Sub
Public Function UltDiaMes(xmes As Long, xano As Long) As Integer
    If xmes = 1 Then
        UltDiaMes = 31
    ElseIf xmes = 2 Then
        If IsDate(zeros(xano, 4) & "/" & "02" & "/" & "29") Then
            UltDiaMes = 29
        Else
            UltDiaMes = 28
        End If
    ElseIf xmes = 3 Then
        UltDiaMes = 31
    ElseIf xmes = 4 Then
        UltDiaMes = 30
    ElseIf xmes = 5 Then
        UltDiaMes = 31
    ElseIf xmes = 6 Then
        UltDiaMes = 30
    ElseIf xmes = 7 Then
        UltDiaMes = 31
    ElseIf xmes = 8 Then
        UltDiaMes = 31
    ElseIf xmes = 9 Then
        UltDiaMes = 30
    ElseIf xmes = 10 Then
        UltDiaMes = 31
    ElseIf xmes = 11 Then
        UltDiaMes = 30
    ElseIf xmes = 12 Then
        UltDiaMes = 31
    End If
End Function
Public Function zeros2(xnumero As String, xqtde As Integer) As String
    zeros2 = String(xqtde - Len(Trim$(xnumero)), "0") & Trim$(xnumero)
End Function
Public Sub travatela(xtela As Form, xTravaDestava As String)
    Dim xcontrol As Control
    For Each xcontrol In xtela.Controls
        If TypeOf xcontrol Is TextBox Then
            If UCase(xTravaDestava) = "D" Then
                xcontrol.BackColor = &HC0FFFF
            ElseIf UCase(xTravaDestava) = "T" Then
                xcontrol.BackColor = &H8000000E
            End If
        End If
    Next
End Sub
Public Function isCNPJ(ByVal pCNPJ As String) As Boolean

    Dim Conta As Integer, soma As Long, Passo As Integer
    Dim Digito1 As Integer, Digito2 As Integer, Flag As Integer

    isCNPJ = False
    pCNPJ = Trim(pCNPJ)

    If Len(pCNPJ) <> 14 Then
        Exit Function
    End If

    For Passo = 5 To 6
        soma = 0
        Flag = Passo
    
        For Conta = 1 To Passo + 7
            soma = soma + (Val(Mid(pCNPJ, Conta, 1)) * Flag)
            Flag = IIf(Flag > 2, Flag - 1, 9)
        Next
    
        soma = soma Mod 11
    
        If Passo = 5 Then Digito1 = IIf(soma > 1, 11 - soma, 0)
        If Passo = 6 Then Digito2 = IIf(soma > 1, 11 - soma, 0)
    Next

    If (Digito1 = Val(Mid(pCNPJ, 13, 1)) And Digito2 = Val(Mid(pCNPJ, 14, 1))) Then
        isCNPJ = True
    End If
    
End Function
Public Function isCPF(ByVal pCPF As String) As Boolean

    Dim Conta As Integer, soma As Integer, Resto As Integer, Passo As Integer
    
    isCPF = False
    pCPF = Trim(pCPF)
    
    If Len(pCPF) <> 11 Then
        Exit Function
    End If
    
    For Passo = 11 To 12
        soma = 0
        For Conta = 1 To Passo - 2
            soma = soma + Val(Mid(pCPF, Conta, 1)) * (Passo - Conta)
        Next
        
        Resto = 11 - (soma - (Int(soma / 11) * 11))
        
        If Resto = 10 Or Resto = 11 Then Resto = 0
        
        If Resto <> Val(Mid(pCPF, Passo - 1, 1)) Then
            Exit Function
        End If
    Next
    isCPF = True
End Function
Public Sub TextMoneyBox_Change(TxtBoxRequired As TextBox)
    If Len(Trim(TxtBoxRequired.Text)) = 0 Then
    TxtBoxRequired.Text = "0,00"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    ElseIf CDbl(TxtBoxRequired.Text) = 0 Then
    TxtBoxRequired.Text = "0,00"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    Else
    TxtBoxRequired.Text = Format((CDbl(SoNumeros(TxtBoxRequired.Text)) / 100), "##,###,##0.00")
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    End If
End Sub
Public Sub TextPesoBox_Change(TxtBoxRequired As TextBox)
    If Len(Trim(TxtBoxRequired.Text)) = 0 Then
    TxtBoxRequired.Text = "0,0"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    ElseIf CDbl(TxtBoxRequired.Text) = 0 Then
    TxtBoxRequired.Text = "0,0"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    Else
    TxtBoxRequired.Text = Format((CDbl(SoNumeros(TxtBoxRequired.Text)) / 10), "##,##0.0")
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    End If
End Sub
Public Function mesnome(nummes As Integer) As String
    If nummes = 1 Then
        mesnome = "JANEIRO"
    ElseIf nummes = 2 Then
        mesnome = "FEVEREIRO"
    ElseIf nummes = 3 Then
        mesnome = "MARCO"
    ElseIf nummes = 4 Then
        mesnome = "ABRIL"
    ElseIf nummes = 5 Then
        mesnome = "MAIO"
    ElseIf nummes = 6 Then
        mesnome = "JUNHO"
    ElseIf nummes = 7 Then
        mesnome = "JULHO"
    ElseIf nummes = 8 Then
        mesnome = "AGOSTO"
    ElseIf nummes = 9 Then
        mesnome = "SETEMBRO"
    ElseIf nummes = 10 Then
        mesnome = "OUTUBRO"
    ElseIf nummes = 11 Then
        mesnome = "NOVEMBRO"
    ElseIf nummes = 12 Then
        mesnome = "DEZEMBRO"
    Else
        MsgBox "Número de Mês Inválido !"
        mesnome = ""
    End If
End Function
Public Function Extenso(ByVal Valor As Double, ByVal MoedaPlural As String, ByVal MoedaSingular As String, Tamanho_String As Integer) As String
    Dim StrValor As String, Negativo As Boolean, Buf As String, Parcial As Integer, Posicao As Integer, Unidades
    Dim Dezenas, Centenas, PotenciasSingular, PotenciasPlural
    
    Negativo = (Valor < 0)
    Valor = Abs(CDec(Valor))
    If Valor Then
        Unidades = Array(vbNullString, "Um", "Dois", "Tres", "Quatro", "Cinco", "Seis", "Sete", "Oito", "Nove", _
                        "Dez", "Onze", "Doze", "Treze", "Quatorze", "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove")
        Dezenas = Array(vbNullString, vbNullString, "Vinte", "Trinta", "Quarenta", "Cinquenta", "Sessenta", "Setenta", _
                        "Oitenta", "Noventa")
        Centenas = Array(vbNullString, "Cento", "Duzentos", "Trezentos", "Quatrocentos", "Quinhentos", _
                        "Seiscentos", "Setecentos", _
                        "Oitocentos", "Novecentos")
        PotenciasSingular = Array(vbNullString, " Mil", " Milhao", " Bilhao", " Trilhao", " Quatrilhao")
        PotenciasPlural = Array(vbNullString, " Mil", " Milhoes", " Bilhoes", " Trilhoes", " Quatrilhoes")
        
        StrValor = Left(Format(Valor, String(18, "0") & ".000"), 18)
        For Posicao = 1 To 18 Step 3
            Parcial = Val(Mid(StrValor, Posicao, 3))
            If Parcial Then
                If Parcial = 1 Then
                    Buf = "Um" & PotenciasSingular((18 - Posicao) \ 3)
                ElseIf Parcial = 100 Then
                    Buf = "Cem" & PotenciasSingular((18 - Posicao) \ 3)
                Else
                    Buf = Centenas(Parcial \ 100)
                    Parcial = Parcial Mod 100
                    If Parcial <> 0 And Buf <> vbNullString Then
                        Buf = Buf & " e "
                    End If
                    If Parcial < 20 Then
                        Buf = Buf & Unidades(Parcial)
                    Else
                        Buf = Buf & Dezenas(Parcial \ 10)
                        Parcial = Parcial Mod 10
                        If Parcial <> 0 And Buf <> vbNullString Then
                            Buf = Buf & " e "
                        End If
                        Buf = Buf & Unidades(Parcial)
                    End If
                    Buf = Buf & PotenciasPlural((18 - Posicao) \ 3)
                End If
                If Buf <> vbNullString Then
                    If Extenso <> vbNullString Then
                    Parcial = Val(Mid(StrValor, Posicao, 3))
                        If Posicao = 16 And (Parcial < 100 Or (Parcial Mod 100) = 0) Then
                            Extenso = Extenso & " e "
                        Else
                            Extenso = Extenso & ", "
                        End If
                    End If
                    Extenso = Extenso & Buf
                End If
            End If
        Next
        If Extenso <> vbNullString Then
            If Negativo Then
                Extenso = "Menos " & Extenso
            End If
            If Int(Valor) = 1 Then
                Extenso = Extenso & " " & MoedaSingular
            Else
                Extenso = Extenso & " " & MoedaPlural
            End If
        End If
        Parcial = Int((Valor - Int(Valor)) * 100 + 0.1)
        If Parcial Then
            Buf = Extenso(Parcial, "Centavos", "Centavo", 0)
            If Extenso <> vbNullString Then
                Extenso = Extenso & " e "
            End If
            Extenso = Extenso & Buf
        End If
        
        If Tamanho_String > 0 Then
            Extenso = Extenso & String(Tamanho_String - Len(Trim$(Extenso)), "*")
        End If
        Extenso = UCase(Extenso)
        
    End If
End Function
Public Function GerarArquivoFaturaBONAGURA()

    Dim xfilial As String, xFatura As String, xEmissao As String, xVencto As String, xvalor As String
    Dim xDesconto As String, xnome As String, xEndereco As String, xTelefone As String, xBairro As String
    Dim xCidade As String, xcep As String, xCnpj As String, xBanco As String, xAgencia As String, xNomeBanco As String
    Dim xSituacao As String, xlinha As String
    Dim xrs As Recordset
    
    'busca Faturas não atualizadas
    If de_informa.rsSel_FaturaEDIContabil.State = 1 Then de_informa.rsSel_FaturaEDIContabil.Close
    de_informa.Sel_FaturaEDIContabil
    
    If de_informa.rsSel_FaturaEDIContabil.RecordCount > 0 Then
    
        'abre arquivo
        Open "C:\INFORMA\CONTABIL\M5FAT" & zeros(Day(datahora("DATA")), 2) & _
                                            zeros(Month(datahora("DATA")), 2) & _
                                            zeros(Hour(datahora("HORA")), 2) & _
                                            zeros(Minute(datahora("HORA")), 2) & ".TXT" For Output As #1
                                                
        Do Until de_informa.rsSel_FaturaEDIContabil.EOF
        
            If de_informa.rsSel_FaturaEDIContabil.Fields("status") = "C" And de_informa.rsSel_FaturaEDIContabil.Fields("at_edi_contabil") = "" Then
                de_informa.rsSel_FaturaEDIContabil.MoveNext
            Else
                xfilial = Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 1, 2)
                
                'FRANKLIN TROQUEI O TAMANHO ABAIXO DE 10 PARA 8
                'xFatura = zeros(CDbl(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 3, 6)), 10)
                xFatura = Space(2) & zeros(CDbl(Mid$(de_informa.rsSel_FaturaEDIContabil.Fields("filialfatura"), 3, 6)), 8)
                '****************fim****************
                
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
        MsgBox "OK ! Arquivo Gerado com sucesso ..."
    
    Else
        MsgBox "Não Há Novas Faturas a Serem Atualizadas !"
        Exit Function
    End If

End Function
    
