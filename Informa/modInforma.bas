Attribute VB_Name = "modInforma"
Option Explicit
Public xusuario As String
Public troca As String
Public xdireitos As String
Public xultimofilial As String
Public xultimoctc As Long
Public xtempoalarme As Integer
Public xamarelo1 As String
Public xamarelo2 As String
Public xbranco As String
Public xstrcon As String
Public xStrconImg As String
'exclusivo para o "checkbox" na flex grid (frmPOD)
Const strChecked = "�" 'Na fonte Wingdings � um Checkbox Checado
Const strUnChecked = "q" 'Na fonte Wingdings � um Checkbox n�o Checado
Public Function transmanif(xfilial As String, xmanifesto As String) As String
    Dim X As Integer
    X = Len(Trim(xmanifesto))
    If X = 1 Then
        xmanifesto = "00000" & Trim(xmanifesto)
    ElseIf X = 2 Then
        xmanifesto = "0000" & Trim(xmanifesto)
    ElseIf X = 3 Then
        xmanifesto = "000" & Trim(xmanifesto)
    ElseIf X = 4 Then
        xmanifesto = "00" & Trim(xmanifesto)
    ElseIf X = 5 Then
        xmanifesto = "0" & Trim(xmanifesto)
    End If
    transmanif = xfilial & Trim(xmanifesto)
End Function
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
Public Function diasprazo(xdataemi As Date, xdataent As Date, xuf As String, xcidade As String, xhora As String, xmodal As String, xFilialIntec As String) As Integer
    Dim xdias1 As Integer, X As Integer, xdt As Date, xdataemi2 As Date, xdiasuteis As Integer
    
    'diferen�a de dias entre entrega e emissao do ctc
    xdias1 = (CDate(Trim$(Str(Year(xdataent))) & "/" & zeros2(Trim$(Str(Month(xdataent))), 2) & "/" & zeros2(Trim$(Str(Day(xdataent))), 2)) - xdataemi)
    
    'se xdias = 0 , emissao = entrega
    If xdias1 = 0 Then
        diasprazo = 0
        Exit Function
    End If
    
    If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
    de_informa.Sel_CadFilial xFilialIntec
        
    xdt = xdataemi 'vari�vel do dia para verifica��o (da emissao at� a entrega)
    
    xdiasuteis = xdias1 'dias �teis (iguala ao total de dias e ir� subtrair os dias n�o �teis)
    
    For X = 0 To xdias1 'la�o do dia 0 at� o dia da entrega
    
        If xdt = xdataemi And xdiasuteis = xdias1 Then   'na data de emissao ...
            'trata dia da emiss�o se � dia �til a cidade da filial emitido
'            If diautil(xdataemi, de_informa.rsSel_CadFilial.Fields("uf"), de_informa.rsSel_CadFilial.Fields("cidade")) = False Then
'                xdiasuteis = xdiasuteis - 1
'            End If
        Else
            'trata fins de semana. Domingo = 1 e S�bado = 7
            If Weekday(xdt) = 1 Or Weekday(xdt) = 7 Then
                If xdt = xdataent Then 'se j� chegou na data da entrega e esta � um feriado n�o abate um dia no c�lculo
                Else
                    xdiasuteis = xdiasuteis - 1
                End If
            Else
                'trata feriados cadastrados
                If de_informa.rsSel_FeriadoNac.State = 1 Then de_informa.rsSel_FeriadoNac.Close
                de_informa.Sel_FeriadoNac Month(xdt), Day(xdt)
                If de_informa.rsSel_FeriadoNac.RecordCount > 0 Then
                    de_informa.rsSel_FeriadoNac.MoveFirst
                    Do Until de_informa.rsSel_FeriadoNac.EOF
                        If de_informa.rsSel_FeriadoNac.Fields("tipo") = "V" Then  'feriado vari�vel
                            If Year(xdt) = de_informa.rsSel_FeriadoNac.Fields("ano") Then 'verif. se bate o ano, pois � feriado vari�vel
                                If xdt = xdataent Then 'se j� chegou na data da entrega e esta � um feriado n�o abate um dia no c�lculo
                                Else
                                    xdiasuteis = xdiasuteis - 1
                                End If
                            End If
                        Else 'feriado fixo, nao verif. o ano pois todo ano � a mesma data
                                If xdt = xdataent Then 'se j� chegou na data da entrega e esta � um feriado n�o abate um dia no c�lculo
                                Else
                                    xdiasuteis = xdiasuteis - 1
                                End If
                        End If
                        de_informa.rsSel_FeriadoNac.MoveNext
                    Loop
                End If
            End If
        End If
        xdt = xdt + 1
    Next
    
    If xdiasuteis < 0 Then xdiasuteis = 0
    
    If xdiasuteis = 0 And xdataemi <> xdataent Then  'se no c�lculo ficou ZERO dias por�m a emiss�o foi em um
        xdiasuteis = 1                               'dia diferente da entrega ent�o diasuteis = 1
    End If
    
    diasprazo = xdiasuteis

End Function
Public Function buscaprazo2(xuf As String, xcidade As String, xtab As String, xmodal As String) As String
    'busca se h� prazo por UF/Cidade espec�fico
    If de_informa.rsSel_CadPrazoCidade.State = 1 Then de_informa.rsSel_CadPrazoCidade.Close
    de_informa.Sel_CadPrazoCidade xtab, Mid$(xmodal, 1, 1), xuf, xcidade
    
    If de_informa.rsSel_CadPrazoCidade.RecordCount > 0 Then
        buscaprazo2 = zeros2(Trim$(Str(de_informa.rsSel_CadPrazoCidade.Fields("prazo"))), 2) & "-" & de_informa.rsSel_CadPrazoCidade.Fields("hscorte")
    Else
        'n�o encontrou por cidade, busca por UF
        If de_informa.rsSel_CadCidade.State = 1 Then de_informa.rsSel_CadCidade.Close
        de_informa.Sel_CadCidade xuf, xcidade
        'busca a tab de prazo por UF
        If de_informa.rsSel_PrazoUF.State = 1 Then de_informa.rsSel_PrazoUF.Close
        de_informa.Sel_PrazoUF xtab, Mid$(xmodal, 1, 1), xuf
        
        If de_informa.rsSel_PrazoUF.RecordCount = 0 Then
            buscaprazo2 = "00-00:00"
        Else
            If de_informa.rsSel_CadCidade.RecordCount = 0 Then  'cidade n�o encontrada trata como interior
                buscaprazo2 = zeros2(Trim$(Str(de_informa.rsSel_PrazoUF.Fields("prazo_int"))), 2) & "-" & de_informa.rsSel_PrazoUF.Fields("hscorte")
            Else
                If de_informa.rsSel_CadCidade.Fields("cim") = "C" Then
                    buscaprazo2 = zeros2(Trim$(Str(de_informa.rsSel_PrazoUF.Fields("prazo_cap"))), 2) & "-" & de_informa.rsSel_PrazoUF.Fields("hscorte")
                Else
                    buscaprazo2 = zeros2(Trim$(Str(de_informa.rsSel_PrazoUF.Fields("prazo_int"))), 2) & "-" & de_informa.rsSel_PrazoUF.Fields("hscorte")
                End If
            End If
        End If
    End If
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
Public Sub Tempo(xtempo As Integer)
Dim xtimer As Long
xtimer = Timer()
Do While True
    If Timer() >= xtimer + xtempo Then
        Exit Sub
    End If
Loop
End Sub
Public Function prev_entr(xdataemi As Date, xuf As String, xcidade As String, xprazo As Integer, xmodal As String, xhora As String, xFilialIntec As String) As Date
    Dim xDutil As Integer, xdutilSN As String, xdt As Date, xDutilContr As Integer
    
    xdt = xdataemi 'vari�vel do dia para verifica��o (da emissao at� a entrega)
    
    If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
    de_informa.Sel_CadFilial xFilialIntec
    
    'verifica se � dia �til na data da emiss�o na cidade da filial emitida
    If diautil(xdataemi, de_informa.rsSel_CadFilial.Fields("uf"), de_informa.rsSel_CadFilial.Fields("cidade")) = False Then
        xDutil = 0  'varialvel que conta o dia util da entrega
    Else
        xDutil = -1  'varialvel que conta o dia util da entrega
    End If
    
    xDutilContr = xDutil
  
    Do While True
        xdutilSN = "S"
        'trata fins de semana. Domingo = 1 e S�bado = 7
        If Weekday(xdt) = 1 Or Weekday(xdt) = 7 Then
            xdutilSN = "N"
        Else
            If xDutil = xDutilContr Then   'dia da emiss�o ou ainda n�o houve o primeiro dia �til, verifica se � dia �til (todos feriados: nacionais, estaduais, locais)
                If diautil(xdt, de_informa.rsSel_CadFilial.Fields("uf"), de_informa.rsSel_CadFilial.Fields("cidade")) = False Then
                    xdutilSN = "N"
                End If
            Else     'verificando dias posteriores a emiss�o: verifica somente feriados estaduais
                'trata feriados NACIONAIS cadastrados
                If de_informa.rsSel_FeriadoNac.State = 1 Then de_informa.rsSel_FeriadoNac.Close
                de_informa.Sel_FeriadoNac Month(xdt), Day(xdt)
                If de_informa.rsSel_FeriadoNac.RecordCount > 0 Then
                    de_informa.rsSel_FeriadoNac.MoveFirst
                    Do Until de_informa.rsSel_FeriadoNac.EOF
                        If de_informa.rsSel_FeriadoNac.Fields("tipo") = "V" Then  'feriado vari�vel
                            If Year(xdt) = de_informa.rsSel_FeriadoNac.Fields("ano") Then 'verif. se bate o ano, pois � feriado vari�vel
                                xdutilSN = "N"
                            End If
                        Else 'feriado fixo, nao verif. o ano pois todo ano � a mesma data
                            xdutilSN = "N"
                        End If
                        de_informa.rsSel_FeriadoNac.MoveNext
                    Loop
                End If
            End If
        End If
        If xdutilSN = "S" Then
            If xprazo = 0 Then Exit Do
            xDutil = xDutil + 1
            If xDutil = xprazo Then Exit Do
            xdt = xdt + 1
        Else
            xdt = xdt + 1
        End If
    Loop
    
    Do While True
        'verifica se � dia �til na data da previs�o de entrega na cidade destino
        If diautil(xdt, xuf, xcidade) = True Then
            Exit Do
        Else
            xdt = xdt + 1
        End If
    Loop
        
    prev_entr = xdt
    
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
Public Sub rel_arquivo()
    Dim xColuna As Integer, xlinha As Integer, xcontctc As Integer, xnumprot As Integer, xcopias As Single, xfil As String
    xfil = InputBox("Escolha a Filial para Gera��o do Protocolo (01, 02, 03, 04, 05, 06, 11 ou % para todas) ?", "Escolha a Filial")
    If xfil <> "%" Then xfil = xfil & "%"
    If de_informa.rssel_RelCtcArquivo.State = 1 Then de_informa.rssel_RelCtcArquivo.Close
    de_informa.sel_RelCtcArquivo xusuario, xfil 'busca CTCs f�sicos baixados por este usu�rio com REL_ARQUIVO = "N"
    If de_informa.rssel_RelCtcArquivo.RecordCount < 1 Then
        MsgBox "N�o h� Dados � serem Impressos !"
        Exit Sub
    Else
    
      de_informa.cn_informa.BeginTrans
        If de_informa.rsSel_RelArqNumero.State = 1 Then de_informa.rsSel_RelArqNumero.Close
        'busca o pr�ximo n�mero do protocolo
        de_informa.Sel_RelArqNumero
        'atualiza o n�mero do pr�ximo protocolo
        de_informa.alt_RelArqNumMais1 Val(de_informa.rsSel_RelArqNumero("ctrrelprotocolo")) + 1
        xnumprot = de_informa.rsSel_RelArqNumero.Fields("ctrrelprotocolo")
        Do Until de_informa.rssel_RelCtcArquivo.EOF
          'atualiza o arquivo com REL_ARQUIVO = "S" + n�mero + Data
          de_informa.alt_RelArquivoSim datahora("data"), xnumprot, de_informa.rssel_RelCtcArquivo("codigo")
          de_informa.rssel_RelCtcArquivo.MoveNext
        Loop
      de_informa.cn_informa.CommitTrans
      
      MsgBox "Ok ! Gerado o Protocolo N�mmero " & Trim$(Str(xnumprot)) & ". Para Enviar Para Impress�o Clique em OK.", vbInformation
      
      For xcopias = 1 To 1
        de_informa.rssel_RelCtcArquivo.MoveFirst
        xcontctc = 0
        xColuna = 1
        xlinha = 0
        Do Until de_informa.rssel_RelCtcArquivo.EOF
            xcontctc = xcontctc + 1   'contador de quantidade
            If xlinha = 0 And xColuna = 1 Then   'identifica inicio da p�gina/cabe��rio
                Printer.Print
                Printer.Print
                Printer.FontSize = 12
                Printer.FontBold = True
                Printer.FontUnderline = True
                Printer.Print Spc(5); "INTEC TRANSPORTES"
                Printer.FontUnderline = False
                Printer.Print
                Printer.Print Spc(5); "RELAT�RIO DE CTCs F�SICOS BAIXADOS"
                Printer.Print Spc(5); "DOCUMENTOS DESPACHADOS PARA SETOR DE ARQUIVO EM " & datahora("data")
                Printer.Print Spc(5); "USU�RIO / DIGITADOR: " & xusuario
                Printer.Print Spc(5); "PROTOCOLO N�MERO: " & String(6 - Len(Trim$(Str(xnumprot))), "0") & Trim$(Str(xnumprot))
                Printer.FontStrikethru = True
                Printer.Print Spc(5); String(140, " ")
                Printer.FontSize = 10
                Printer.FontStrikethru = False
                Printer.FontBold = False
                Printer.FontUnderline = False
            End If
            'impress�o por 6 colunas
            If xColuna = 1 Then
                Printer.Print Spc(6); Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 1, 2) & "-" & _
                Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 3, 8);
                xColuna = xColuna + 1
                xlinha = xlinha + 1
            ElseIf xColuna = 2 Then
                Printer.Print Spc(3); Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 1, 2) & "-" & _
                Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 3, 8);
                xColuna = xColuna + 1
            ElseIf xColuna = 3 Then
                Printer.Print Spc(3); Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 1, 2) & "-" & _
                Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 3, 8);
                xColuna = xColuna + 1
            ElseIf xColuna = 4 Then
                Printer.Print Spc(3); Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 1, 2) & "-" & _
                Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 3, 8);
                xColuna = xColuna + 1
            ElseIf xColuna = 5 Then
                Printer.Print Spc(3); Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 1, 2) & "-" & _
                Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 3, 8);
                xColuna = xColuna + 1
            ElseIf xColuna = 6 Then
                Printer.Print Spc(3); Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 1, 2) & "-" & _
                Mid(de_informa.rssel_RelCtcArquivo.Fields("filialctc"), 3, 8);
                'na coluna 6 volta para a coluna 1
                xColuna = 1
                Printer.Print
                'e se a linha for = 26 e n�o for �ltimo CTC ...
                If xlinha = 26 And de_informa.rssel_RelCtcArquivo.RecordCount <> xcontctc Then
                    xlinha = 0
                    Printer.FontSize = 12
                    Printer.FontBold = True
                    Printer.FontStrikethru = True
                    Printer.Print Spc(5); String(140, " ")
                    Printer.FontSize = 10
                    Printer.FontBold = False
                    Printer.FontStrikethru = False
                    Printer.Print
                    Printer.Print
                    Printer.Print Spc(7); "Visto Confer�ncia: ______________"
                    Printer.NewPage
                Else
                    Printer.Print
                End If
            End If
            de_informa.rssel_RelCtcArquivo.MoveNext
        Loop
        'final do relat�rio, fecha com rodap�
        Printer.Print
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.FontStrikethru = True
        Printer.Print Spc(5); String(140, " ")
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontStrikethru = False
        Printer.Print
        Printer.Print Spc(7); "Quantidade de CTCs: "; xcontctc
        Printer.Print Spc(7); "Data: "; datahora("data"); Space(35); "Assinatura Confer�ncia: ______________________"
        Printer.NewPage
      Next
        Printer.EndDoc   'finaliza spool da impress�o
        
        'LOG DE USU�RIO
        de_informa.ins_LogUsuario "IMPRESS�O", xusuario, "IMPRESS�O DO PROTOCOLO: " & Trim$(Str(xnumprot))
        
        MsgBox "RELAT�RIO ENVIADO PARA IMPRESS�O ! PROTOCOLO N�M: " & String(6 - Len(Trim$(Str(xnumprot))), "0") & Trim$(Str(xnumprot))
    End If
End Sub
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
        MsgBox "ERRO ! Par�metro Inv�lido na Fun��o DATAHORA() !", vbCritical, "ERRO de SISTEMA"
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

Public Function PriMaiuscula(Texto) As String
Dim xMaiuscula As String
Dim xTexto2 As String
Dim X As Integer

Texto = LCase(Texto)
xMaiuscula = "SIM"
xTexto2 = ""

        For X = 1 To Len((Texto)) Step 1
           If xMaiuscula = "SIM" Then
            xTexto2 = xTexto2 & UCase(Mid((Texto), X, 1))
            Else
            xTexto2 = xTexto2 & Mid((Texto), X, 1)
            End If

            If Mid((Texto), X, 1) = " " Or Mid((Texto), X, 1) = "." Or Mid((Texto), X, 1) = "/" Or Mid((Texto), X, 1) = "\" Or Mid((Texto), X, 1) = ";" Or Mid((Texto), X, 1) = ":" Or Mid((Texto), X, 1) = "_" Or Mid((Texto), X, 1) = "&" Or Mid((Texto), X, 1) = "-" Then
                If Mid((Texto), X, 1) = " " Then
                    If Mid((Texto), X, 4) = " do " Then
                    xMaiuscula = "NAO"
                    ElseIf Mid((Texto), X, 4) = " da " Then
                    xMaiuscula = "NAO"
                    ElseIf Mid((Texto), X, 4) = " de " Then
                    xMaiuscula = "NAO"
                    ElseIf Mid((Texto), X, 5) = " das " Then
                    xMaiuscula = "NAO"
                    ElseIf Mid((Texto), X, 5) = " dos " Then
                    xMaiuscula = "NAO"
                    Else
                    xMaiuscula = "SIM"
                    End If
                Else
                xMaiuscula = "SIM"
                End If
            Else
            xMaiuscula = "NAO"
            End If
        Next
        
PriMaiuscula = xTexto2
End Function
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
    'xDATA_INICIAL = xMskEdBox.Text
    End If
    'xDATA_INICIAL = xMskEdBox.Text
End Sub

Public Sub TravaTela(xtela As Form, xTravaDestava As String)
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
Public Function zeros2(xnumero As String, xqtde As Integer) As String
    zeros2 = String(xqtde - Len(Trim$(xnumero)), "0") & Trim$(xnumero)
End Function
Public Function diautil(xdata As Date, xuf As String, xcidade As String) As Boolean
    If Weekday(xdata) = 1 Or Weekday(xdata) = 7 Then
        diautil = False
        Exit Function
    Else
        If de_informa.rsSel_Feriado.State = 1 Then de_informa.rsSel_Feriado.Close
        de_informa.Sel_Feriado Month(xdata), Day(xdata)
        If de_informa.rsSel_Feriado.RecordCount > 0 Then
            de_informa.rsSel_Feriado.MoveFirst
            Do Until de_informa.rsSel_Feriado.EOF
                If de_informa.rsSel_Feriado.Fields("uf") = "BR" Then 'feriado nacional
                    If de_informa.rsSel_Feriado.Fields("tipo") = "V" Then  'feriado vari�vel
                        If Year(xdata) = de_informa.rsSel_Feriado.Fields("ano") Then 'verif. se bate o ano, pois � feriado vari�vel
                            diautil = False
                            Exit Function
                        End If
                    Else 'feriado fixo, nao verif. o ano pois todo ano � a mesma data
                        diautil = False
                        Exit Function
                    End If
                ElseIf de_informa.rsSel_Feriado.Fields("uf") <> "BR" _
                And de_informa.rsSel_Feriado.Fields("cidade") = "" Then 'feriado estadual
                    If xuf = de_informa.rsSel_Feriado.Fields("uf") Then
                        diautil = False
                        Exit Function
                    End If
                ElseIf de_informa.rsSel_Feriado.Fields("uf") <> "BR" _
                And de_informa.rsSel_Feriado.Fields("cidade") <> "" Then 'feriado local/municipal
                    If xuf = de_informa.rsSel_Feriado.Fields("uf") _
                    And xcidade = de_informa.rsSel_Feriado.Fields("cidade") Then
                        diautil = False
                        Exit Function
                    End If
                End If
                de_informa.rsSel_Feriado.MoveNext
            Loop
            
        End If
    End If
    
    diautil = True

End Function
'exclusivo para o "checkbox" na flex grid (frmPOD)
Public Function VerificaCheck(iRow As Integer, iCol As Integer, Formulario As Form)
    
    With Formulario.MSFlexGrid1 'Determinal o Formulario Que Esta o CheckBox e Seta o Msflexgrid
        If .TextMatrix(iRow, iCol) = strUnChecked Then 'se Check n�o estiver Marcado
            .TextMatrix(iRow, iCol) = strChecked 'Marca o Checkbox
        ElseIf .TextMatrix(iRow, iCol) = strChecked Then 'Se Estiver Marcado
            .TextMatrix(iRow, iCol) = strUnChecked ' Desmarca o Checkbox
        Else 'Caso N�o Possua um CheckBox
        End If
    End With
    
End Function
'exclusivo para o "checkbox" na flex grid (frmPOD)
Public Function ColocaCheck(iRow As Integer, iCol As Integer, Formulario As Form)
  With Formulario.MSFlexGrid1 'Determinal o Formulario Que Esta o CheckBox e Seta o Msflexgrid
    .Row = iRow 'Pega a Linha Para inserir o CheckBox
    .Col = iCol 'Pega a Coluna Para Inserir O CheckBox
    .CellFontName = "Wingdings" 'Muda A Fonte da Determinada Celula do Flexgrid Para da O efeito de CheckBox
    .CellFontSize = 14 'Muda o Tamanho da Fonte daquela Celula do Flexgrid
    .CellAlignment = flexAlignCenterCenter 'Posiciona o CheckBox no Centro da Celula
    .Text = strUnChecked 'Usa a Constante Para Inserir o Caracter n�o unChecked
  End With
End Function
