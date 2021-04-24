Attribute VB_Name = "ModEmissao"
Public xUsuario As String
Public xUsuarioIMP As String
Public xDataIMP As String
Public xHoraIMP As String
Public StringDireitos As String
Public xForm As Form
Public xAmarelo As Long
Public xBranco As Long
Public xAzul As Long
Public xPreto As Long
Public xLaranja As Long
Public xCinzaClaro As Long
Public Leave, LeaveSub As Boolean
Public Acao As String
Public SETIMPImpressoraPadrao As String
Public xACOMP As Boolean


Public AUXCanc As String


Public Function SoNumero(xNumero As String) As String
Dim xCont As Integer, xTextoAux As String
xTextoAux = ""
    If Len(Trim$(xNumero)) > 0 Then
        For xCont = 1 To Len(Trim$(xNumero))
            If IsNumeric(Mid(xNumero, xCont, 1)) = True Then
            xTextoAux = xTextoAux & Mid(xNumero, xCont, 1)
            End If
        Next
    End If
SoNumero = xTextoAux
End Function

Public Function SemAcento(Texto As String) As String
Dim Step, TxtCode As Long

    For Step = 1 To Len(Texto)
    
    TxtCode = Asc(Mid(Texto, Step, 1))
    
    If TxtCode >= 192 And TxtCode <= 197 Then TxtCode = 65
    If TxtCode = 199 Then TxtCode = 67
    If TxtCode >= 200 And TxtCode <= 203 Then TxtCode = 69
    If TxtCode >= 204 And TxtCode <= 207 Then TxtCode = 73
    If TxtCode = 209 Then TxtCode = 78
    If TxtCode >= 210 And TxtCode <= 214 Then TxtCode = 79
    If TxtCode >= 217 And TxtCode <= 220 Then TxtCode = 85
    
    If TxtCode >= 224 And TxtCode <= 229 Then TxtCode = 97
    If TxtCode = 231 Then TxtCode = 99
    If TxtCode >= 232 And TxtCode <= 235 Then TxtCode = 101
    If TxtCode >= 236 And TxtCode <= 239 Then TxtCode = 105
    If TxtCode = 241 Then TxtCode = 110
    If TxtCode >= 242 And TxtCode <= 246 Then TxtCode = 111
    If TxtCode >= 249 And TxtCode <= 252 Then TxtCode = 117

    Mid(Texto, Step, 1) = Chr(TxtCode)
    Next
SemAcento = Texto
End Function


Public Function SemPonto(Texto As String) As Long
Dim Cont As Integer
Dim TextoAux As String
TextoAux = ""
    For Cont = 1 To Len(Texto)
        If Mid(Texto, Cont, 1) <> "." And Mid(Texto, Cont, 1) <> "," And Mid(Texto, Cont, 1) <> "%" Then
        TextoAux = TextoAux & Mid(Texto, Cont, 1)
        End If
    Next
    
    If Len(TextoAux) <= 9 Then
    SemPonto = Val(TextoAux)
    Else
    SemPonto = 0
    End If
End Function

Public Sub TextMoneyBox_KeyPress(KeyAsciiRequired As Integer)
    If KeyAsciiRequired < 48 Or KeyAsciiRequired > 57 Then
        If KeyAsciiRequired = 13 Then
        SendKeys "{TAB}"
        Else
            If KeyAsciiRequired <> 8 Then
            KeyAsciiRequired = 0
            End If
        End If
    End If
End Sub

Public Sub TextMoneyBox_GotFocus(TxtBoxRequired As TextBox)
If Trim(TxtBoxRequired.Text) = "" Then TxtBoxRequired.Text = "0,00"
TxtBoxRequired.SelStart = 0
TxtBoxRequired.SelLength = Len(TxtBoxRequired.Text) + 1
End Sub

Public Sub TextMoneyBox_Change(TxtBoxRequired As TextBox)
    If Len(Trim(TxtBoxRequired.Text)) = 0 Then
    TxtBoxRequired.Text = "0.00"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    ElseIf CDbl(TxtBoxRequired.Text) = 0 Then
    TxtBoxRequired.Text = "0.00"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    Else
    TxtBoxRequired.Text = Format((SemPonto(TxtBoxRequired.Text) / 100), "###,##0.00")
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    End If
End Sub

Public Sub Date_MskEdBox_GotFocus(xMskEdBox As MaskEdBox)
xMskEdBox.Mask = "##/##/####"
xMskEdBox.SelStart = 0
xMskEdBox.SelLength = 100
End Sub

Public Sub Date_MskEdBox_LostFocus(xMskEdBox As MaskEdBox)
    If Mid(Trim(xMskEdBox.Text), 9, 2) = "__" And IsDate(Mid(Trim(xMskEdBox.Text), 1, 8)) = True Then
    xMskEdBox.Text = Mid(xMskEdBox.Text, 1, 6) & "20" & Mid(xMskEdBox.Text, 7, 2)
    End If
    
    If Not IsDate(xMskEdBox) And xMskEdBox.Text <> "__/__/____" Then
    xMskEdBox.Mask = "##/##/####"
    xMskEdBox.Text = "__/__/____"
    xMskEdBox.SetFocus
    End If
    If xMskEdBox.Text = "__/__/____" Or xMskEdBox.Text = "" Then
    xMskEdBox.Mask = ""
    xMskEdBox.Text = ""
    xDATA_INICIAL = xMskEdBox.Text
    End If
    xDATA_INICIAL = xMskEdBox.Text
End Sub

Public Function PriMaiuscula(Texto) As String
Texto = LCase(Texto)
xmaiuscula = "SIM"
xtexto2 = ""


        For X = 1 To Len(Trim(Texto)) Step 1
           If xmaiuscula = "SIM" Then
            xtexto2 = xtexto2 & UCase(Mid(Trim(Texto), X, 1))
            Else
            xtexto2 = xtexto2 & Mid(Trim(Texto), X, 1)
            End If

            If Mid(Trim(Texto), X, 1) = " " Or Mid(Trim(Texto), X, 1) = "." Or Mid(Trim(Texto), X, 1) = "/" Or Mid(Trim(Texto), X, 1) = "\" Or Mid(Trim(Texto), X, 1) = ";" Or Mid(Trim(Texto), X, 1) = ":" Or Mid(Trim(Texto), X, 1) = "_" Or Mid(Trim(Texto), X, 1) = "&" Or Mid(Trim(Texto), X, 1) = "-" Then
                If Mid(Trim(Texto), X, 1) = " " Then
                    If Mid(Trim(Texto), X, 4) = " do " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 4) = " da " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 4) = " de " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 5) = " das " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 5) = " dos " Then
                    xmaiuscula = "NAO"
                    Else
                    xmaiuscula = "SIM"
                    End If
                Else
                xmaiuscula = "SIM"
                End If
            Else
            xmaiuscula = "NAO"
            End If
        Next
        
PriMaiuscula = xtexto2
End Function

Public Function EspCaracat(Texto As String) As Boolean
xcaracteresespeciais = "|\<,>.:;?/°^~}]º{[ª`´Çç'!¹@²#³$£%¢¨¬&*()_-+=§"
    For xpos = 1 To Len(xcaracteresespeciais)
        If Texto = Mid(xcaracteresespeciais, xpos, 1) Then
        EspCaracat = True
        xpos = Len(xcaracteresespeciais)
        Else
        EspCaracat = False
        End If
    Next
End Function

Public Function ExisteArquivo(xCaminhoArquivo As String) As Boolean

ExisteArquivo = False

    If Len(Dir(xCaminhoArquivo, vbDirectory)) <> 0 Then
    ExisteArquivo = True
    End If

End Function

Public Sub OrdenaListBox(xListBox As ListBox)
xListBox.Visible = False
DoEvents
    For i = 0 To xListBox.ListCount - 1
        For j = 0 To xListBox.ListCount - 2
            If xListBox.List(i) < xListBox.List(j) Then
            xAUX = xListBox.List(j)
            xListBox.List(j) = xListBox.List(i)
            xListBox.List(i) = xAUX
            End If
        Next
    Next
xListBox.Visible = True
DoEvents
End Sub

Public Sub TransfereItemDeListBox(List_Origem As ListBox, List_Destino As ListBox)
List_Origem.Visible = False
List_Destino.Visible = False
DoEvents
    For X = 0 To List_Origem.ListCount - 1
        If List_Origem.Selected(X) = True Then
        List_Destino.AddItem List_Origem.List(X)
        End If
    Next
    
X = 0
    Do While True
        
        If X > List_Origem.ListCount - 1 Then
        Exit Do
        End If
        
        If List_Origem.Selected(X) = True Then
        List_Origem.RemoveItem (X)
        X = X - 1
        End If
    X = X + 1
    Loop
List_Origem.Visible = True
List_Destino.Visible = True
DoEvents
End Sub

Public Sub TextPesoBox_KeyPress(KeyAsciiRequired As Integer)
    If KeyAsciiRequired < 48 Or KeyAsciiRequired > 57 Then
        If KeyAsciiRequired = 13 Then
        SendKeys "{TAB}"
        Else
            If KeyAsciiRequired <> 8 Then
            KeyAsciiRequired = 0
            End If
        End If
    End If
End Sub

Public Sub TextPesoBox_GotFocus(TxtBoxRequired As TextBox)
If Trim(TxtBoxRequired.Text) = "" Then TxtBoxRequired.Text = "0,0"
TxtBoxRequired.SelStart = 0
TxtBoxRequired.SelLength = Len(TxtBoxRequired.Text) + 1
End Sub

Public Sub TextPesoBox_Change(TxtBoxRequired As TextBox)
    If Len(Trim(TxtBoxRequired.Text)) = 0 Then
    TxtBoxRequired.Text = "0.0"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    ElseIf CDbl(TxtBoxRequired.Text) = 0 Then
    TxtBoxRequired.Text = "0.0"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    Else
    TxtBoxRequired.Text = Format((SemPonto(TxtBoxRequired.Text) / 10), "###,##0.0")
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    End If
End Sub

Public Function JaExisteTabelaOficial(Sigla_CiaAerea As String, Sigla_Origem As String) As Long
JaExisteTabelaOficial = 0

    If de_informa.rsVerificacaoTabOFICIAL.State = 1 Then de_informa.rsVerificacaoTabOFICIAL.Close
    de_informa.VerificacaoTabOFICIAL Sigla_CiaAerea, Sigla_Origem
    
    If de_informa.rsVerificacaoTabOFICIAL.RecordCount > 0 Then
    JaExisteTabelaOficial = de_informa.rsVerificacaoTabOFICIAL.Fields("codtab")
    End If
End Function

Public Function JaExisteTabelaEspecifica(Sigla_CiaAerea As String, Sigla_Origem As String, CGC_Cliente As String) As Long
JaExisteTabelaEspecifica = 0

    If de_informa.rsVerificacaoTabESPECIFICA.State = 1 Then de_informa.rsVerificacaoTabESPECIFICA.Close
    de_informa.VerificacaoTabESPECIFICA Sigla_CiaAerea, Sigla_Origem, CGC_Cliente
    
    If de_informa.rsVerificacaoTabESPECIFICA.RecordCount > 0 Then
    JaExisteTabelaEspecifica = de_informa.rsVerificacaoTabESPECIFICA.Fields("codtab")
    End If
End Function


Public Function DataHora(xparametro As String) As Variant
Dim xretorno As Variant
If UCase(xparametro) = "DATA" Then
    If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
    de_informa.Sel_DataServidor
    xretorno = CDate(Trim$(Str(Year(de_informa.rsSel_DataServidor.Fields("agora")))) & "/" & _
Trim$(Str(Month(de_informa.rsSel_DataServidor.Fields("agora")))) & "/" & _
Trim$(Str(Day(de_informa.rsSel_DataServidor.Fields("agora")))))
    DataHora = xretorno
ElseIf UCase(xparametro) = "HORA" Then
    If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
    de_informa.Sel_DataServidor
    xretorno = Zeros(Hour(de_informa.rsSel_DataServidor.Fields("agora")), 2) & ":" & _
               Zeros(Minute(de_informa.rsSel_DataServidor.Fields("agora")), 2) & ":" & _
               Zeros(Second(de_informa.rsSel_DataServidor.Fields("agora")), 2)
    DataHora = xretorno
ElseIf UCase(xparametro) = "DATAHORA" Then
    If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
    de_informa.Sel_DataServidor
    xretorno = Trim$(de_informa.rsSel_DataServidor.Fields("agora"))
    DataHora = xretorno
Else
    MsgBox "ERRO! Parâmetro Inválido na Função DATAHORA()!", vbCritical, "ERRO de SISTEMA"
    DataHora = ""
End If

End Function

Public Function Zeros(xNumero As Long, xQtde As Integer) As String
Zeros = String(xQtde - Len(Trim$(Str(xNumero))), "0") & Trim$(Str(xNumero))
End Function

Public Sub LimpaTela(xTela As Form)
    Dim xmask As String
    Dim xControl As Control
    For Each xControl In xTela.Controls
        If TypeOf xControl Is TextBox Then
            xControl.Text = ""
        ElseIf TypeOf xControl Is Label Then
            If xControl.BorderStyle = 1 Then
                xControl.Caption = ""
            End If
        ElseIf TypeOf xControl Is MaskEdBox Then
            xControl.Mask = ""
            xControl.Text = ""
        End If
    Next
End Sub

Public Sub LimpaFrame(xTela As Form, xFrameCaption As String)
    Dim xmask As String
    Dim xControl As Control
    For Each xControl In xTela
        If TypeOf xControl Is TextBox Then
            If xControl.Container = xFrameCaption Then
            xControl.Text = ""
            End If
        ElseIf TypeOf xControl Is Label Then
            If xControl.Container = xFrameCaption Then
                If xControl.BorderStyle = 1 Then
                    xControl.Caption = ""
                End If
            End If
        ElseIf TypeOf xControl Is MaskEdBox Then
            If xControl.Container = xFrameCaption Then
            xControl.Mask = ""
            xControl.Text = ""
            End If
        End If
    Next
End Sub

Public Sub TravaFrame(xTela As Form, xFrame As Frame, Tipo As Integer)
    Dim xControl As Control
    For Each xControl In xTela
        If TypeOf xControl Is Frame And xControl <> xFrame And Tipo = 0 Then
        xControl.Enabled = False
        ElseIf TypeOf xControl Is Frame And xControl <> xFrame And Tipo = 1 Then
        xControl.Enabled = True
        End If
    Next
End Sub
Public Function isCNPJ(ByVal pCNPJ As String) As Boolean

    Dim Conta As Integer, Soma As Long, Passo As Integer
    Dim Digito1 As Integer, Digito2 As Integer, Flag As Integer

    isCNPJ = False
    pCNPJ = Trim(pCNPJ)

    If Len(pCNPJ) <> 14 Then
        Exit Function
    End If

    For Passo = 5 To 6
        Soma = 0
        Flag = Passo

        For Conta = 1 To Passo + 7
            Soma = Soma + (Val(Mid(pCNPJ, Conta, 1)) * Flag)
            Flag = IIf(Flag > 2, Flag - 1, 9)
        Next

        Soma = Soma Mod 11

        If Passo = 5 Then Digito1 = IIf(Soma > 1, 11 - Soma, 0)
        If Passo = 6 Then Digito2 = IIf(Soma > 1, 11 - Soma, 0)
    Next

    If (Digito1 = Val(Mid(pCNPJ, 13, 1)) And Digito2 = Val(Mid(pCNPJ, 14, 1))) Then
        isCNPJ = True
    End If

End Function


Public Function isCPF(ByVal pCPF As String) As Boolean

    Dim Conta As Integer, Soma As Integer, Resto As Integer, Passo As Integer

    isCPF = False
    pCPF = Trim(pCPF)

    If Len(pCPF) <> 11 Then
        Exit Function
    End If

    For Passo = 11 To 12
        Soma = 0
        For Conta = 1 To Passo - 2
            Soma = Soma + Val(Mid(pCPF, Conta, 1)) * (Passo - Conta)
        Next

        Resto = 11 - (Soma - (Int(Soma / 11) * 11))

        If Resto = 10 Or Resto = 11 Then Resto = 0

        If Resto <> Val(Mid(pCPF, Passo - 1, 1)) Then
            Exit Function
        End If
    Next
    isCPF = True
End Function

Public Function TransCodAWB(xxFilial As String, xxCia As String, xxAWB As String, xxDig As String) As String
TransCodAWB = String(2 - Len(Trim(Str(Val(Mid(xxFilial, 1, 2))))), "0") & Trim(Str(Val(Mid(xxFilial, 1, 2)))) & UCase(Trim(Mid(xxCia, 1, 2))) & String(10 - Len(Trim(Str(Val(Mid(xxAWB, 1, 10))))), "0") & Trim(Str(Val(Mid(xxAWB, 1, 10)))) & Mid(Trim(Str(Val(xDig))), 1, 1)
End Function

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

Public Function NomeMes(NumeroMes As Integer) As String
NomeMes = ""
If NumeroMes = 1 Then NomeMes = "Janeiro"
If NumeroMes = 2 Then NomeMes = "Fevereiro"
If NumeroMes = 3 Then NomeMes = "Março"
If NumeroMes = 4 Then NomeMes = "Abril"
If NumeroMes = 5 Then NomeMes = "Maio"
If NumeroMes = 6 Then NomeMes = "Junho"
If NumeroMes = 7 Then NomeMes = "Julho"
If NumeroMes = 8 Then NomeMes = "Agosto"
If NumeroMes = 9 Then NomeMes = "Setembro"
If NumeroMes = 10 Then NomeMes = "Outubro"
If NumeroMes = 11 Then NomeMes = "Novembro"
If NumeroMes = 12 Then NomeMes = "Dezembro"
End Function

Public Function NumMes(NomedoMes As String) As Integer
NomedoMes = LCase(Trim(NomedoMes))
NumMes = 0
If NomedoMes = "janeiro" Then NumMes = 1
If NomedoMes = "fevereiro" Then NumMes = 2
If NomedoMes = "março" Then NumMes = 3
If NomedoMes = "abril" Then NumMes = 4
If NomedoMes = "maio" Then NumMes = 5
If NomedoMes = "junho" Then NumMes = 6
If NomedoMes = "julho" Then NumMes = 7
If NomedoMes = "agosto" Then NumMes = 8
If NomedoMes = "setembro" Then NumMes = 9
If NomedoMes = "outubro" Then NumMes = 10
If NomedoMes = "novembro" Then NumMes = 11
If NomedoMes = "dezembro" Then NumMes = 12
End Function

Public Sub AtualizaStatusTabelas()
de_informa.cn_informa.BeginTrans
de_informa.UpdateExpiraTabela CDate(DataHora("DATA"))
de_informa.UpdateVigoraTabela CDate(DataHora("DATA"))
de_informa.cn_informa.CommitTrans
End Sub
