VERSION 5.00
Begin VB.Form frmGeraProtCanhotos 
   Caption         =   "Controle de Canhotos para o Cliente (Geração/Impressão Protocolo)"
   ClientHeight    =   5730
   ClientLeft      =   1875
   ClientTop       =   1680
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7485
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Excluir/Incluir Canhotos a um Protocolo Gerado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   7215
      Begin VB.CommandButton cmdGravarIncExcl 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   5280
         TabIndex        =   20
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtExcCanhoto 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtExcCtc 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtExcFilial 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   27
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtIncCanhoto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtIncCtc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtIncFilial 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   22
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtExcProtocolo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtIncProtocolo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optExcluir 
         Caption         =   "Excluir:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optIncluir 
         Caption         =   "Incluir:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Canhoto NF:"
         Height          =   195
         Left            =   5040
         TabIndex        =   29
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Filial-CTC:"
         Height          =   195
         Left            =   2640
         TabIndex        =   26
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Canhoto NF:"
         Height          =   195
         Left            =   5040
         TabIndex        =   24
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filial-CTC:"
         Height          =   195
         Left            =   2640
         TabIndex        =   21
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo:"
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo:"
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gerar / Imprimir o Protocolo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar/Imprimir..."
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtcopias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "3"
         Top             =   1560
         Width           =   375
      End
      Begin VB.OptionButton optGerar 
         Caption         =   "Gerar/Imprimir Protocolo do Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optRegerar 
         Caption         =   "Re-Imprimir o Protocolo número: "
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtRemetCGC 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   14
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtNumProt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdBuscaCli 
         Caption         =   "?"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox chkTodosEstab 
         Caption         =   "Todos Estabelecimentos do Cliente"
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.Label lblRemetNome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   30
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número de Vias:"
         Height          =   195
         Left            =   5280
         TabIndex        =   11
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC:"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmGeraProtCanhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaCli_Click()
    frmBuscaCLI.Caption = "Busca Cliente - Controle de Canhotos de NF"
    frmBuscaCLI.Show 1
End Sub

Private Sub cmdGravarIncExcl_Click()
    If Val(txtExcProtocolo) > 0 And IsNumeric(txtExcProtocolo) = True Then
    Else
        MsgBox "DADOS INVÁLIDOS: NUM. DO PROTOCOLO !"
        txtExcProtocolo.SetFocus
        Exit Sub
    End If
    If Val(txtExcCanhoto) > 0 And IsNumeric(txtExcCanhoto) = True Then
    Else
        MsgBox "DADOS INVÁLIDOS: NUM. DO NF/CANHOTO !"
        txtExcCanhoto.SetFocus
        Exit Sub
    End If

    If de_informa.rsSel_CanhotosProtNF.State = 1 Then de_informa.rsSel_CanhotosProtNF.Close
    de_informa.Sel_CanhotosProtNF Val(txtExcProtocolo), Val(txtExcCanhoto)
    If de_informa.rsSel_CanhotosProtNF.RecordCount > 0 Then
        de_informa.Excl_CanhotosProtNF Val(txtExcProtocolo), Val(txtExcCanhoto)
        MsgBox "OK ! NF/Canhoto Excluído do Protocolo !", vbOKOnly + vbInformation
        txtExcProtocolo.SetFocus
        de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "EXCLUIR CANHOTO " & Trim$(txtExcCanhoto) & " DO PROTOCOLO " & Trim$(txtExcProtocolo)
    Else
        MsgBox "NF/Canhoto não Encontrado para este Protocolo !", vbOKOnly + vbCritical
        txtExcProtocolo.SetFocus
    End If
End Sub

Private Sub CmdProcessar_Click()
    Dim xColuna As Long, xlinha As Long, xcontnf As Long, xnumprot As Long, xcopias As Integer
    
If Not Len(Trim$(txtRemetCGC)) > 0 And optGerar.Value = True Then
    MsgBox "É Necessário Escolher um Cliente Para Gerar os Protocolos de NF !"
    txtRemetCGC.SetFocus
    Exit Sub
End If
    
If optGerar.Value = True Then
    If de_informa.rsSel_GerarCanhotosNF.State = 1 Then de_informa.rsSel_GerarCanhotosNF.Close
    de_informa.Sel_GerarCanhotosNF Trim$(txtRemetCGC) & "%", xusuario
    If de_informa.rsSel_GerarCanhotosNF.RecordCount < 1 Then
        MsgBox "Não Há Dados à Serem Gerados Para Este Cliente neste Usuário!"
        Exit Sub
    Else
        de_informa.rsSel_GerarCanhotosNF.MoveFirst
        
        'busca o próximo número do protocolo
        If de_informa.rsSel_RelArqNumero.State = 1 Then de_informa.rsSel_RelArqNumero.Close
        de_informa.Sel_RelArqNumero
        
        'inicia transaçao
        de_informa.cn_informa.BeginTrans
        
        'atualiza o número do próximo protocolo
        xnumprot = de_informa.rsSel_RelArqNumero.Fields("ctrrelcanhoto")
        de_informa.alt_RelArqNumCanhotoMais1 Val(de_informa.rsSel_RelArqNumero("ctrrelcanhoto")) + 1

        'atualiza o arquivo com o número do protocolo
        Do Until de_informa.rsSel_GerarCanhotosNF.EOF
            'de_informa.alt_RelCanhotoNF xnumprot, de_informa.rsSel_GerarCanhotosNF.Fields("filialctc")
            de_informa.rsSel_GerarCanhotosNF.MoveNext
        Loop
        
        'finaliza transação
        de_informa.cn_informa.CommitTrans
        
        'inicial a impressão
        de_informa.rsSel_GerarCanhotosNF.MoveFirst
        
        For xcopias = 1 To Val(txtcopias)
            de_informa.rsSel_GerarCanhotosNF.MoveFirst
            xcontnf = 0
            xColuna = 1
            xlinha = 0
            Do Until de_informa.rsSel_GerarCanhotosNF.EOF
                xcontnf = xcontnf + 1   'contador de quantidade
                If xlinha = 0 And xColuna = 1 Then   'identifica inicio da página/cabeçário
                    Printer.Print
                    Printer.Print
                    Printer.FontSize = 12
                    Printer.FontBold = True
                    Printer.FontUnderline = True
                    Printer.Print Spc(5); "INTEC TRANSPORTES"
                    Printer.FontUnderline = False
                    Printer.Print
                    Printer.Print Spc(5); "RELATÓRIO DE CANHOTOS DAS NOTAS FISCAIS DEVOLVIDOS PARA O CLIENTE"
                    Printer.Print Spc(5); "DATA: " & datahora("data")
                    Printer.Print Spc(5); "USUÁRIO: " & xusuario
                    Printer.Print Spc(5); "PROTOCOLO NÚMERO: " & zeros(xnumprot, 6)
                    Printer.Print Spc(5); "CLIENTE: " & Trim$(lblRemetNome)
                    Printer.FontStrikethru = True
                    Printer.Print Spc(5); String(140, " ")
                    Printer.FontSize = 10
                    Printer.FontStrikethru = False
                    Printer.FontBold = False
                    Printer.FontUnderline = False
                End If
                'impressão por 6 colunas
                If xColuna = 1 Then
                    Printer.Print Spc(6); zeros(de_informa.rsSel_GerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                    xlinha = xlinha + 1
                ElseIf xColuna = 2 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_GerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 3 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_GerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 4 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_GerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 5 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_GerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 6 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_GerarCanhotosNF.Fields("numnf"), 10);
                    'na coluna 6 volta para a coluna 1
                    xColuna = 1
                    Printer.Print
                    'e se a linha for = 26 e não for último CTC ...
                    If xlinha = 26 And de_informa.rsSel_GerarCanhotosNF.RecordCount <> xcontnf Then
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
                        Printer.Print Spc(7); "Visto Conferência: ______________"
                        Printer.NewPage
                    Else
                        'Printer.Print
                    End If
                End If
                de_informa.rsSel_GerarCanhotosNF.MoveNext
            Loop
            de_informa.rsSel_GerarCanhotosNF.MoveFirst
            Do Until de_informa.rsSel_GerarCanhotosNF.EOF
                de_informa.alt_RelCanhotoNF xnumprot, de_informa.rsSel_GerarCanhotosNF.Fields("filialctc"), de_informa.rsSel_GerarCanhotosNF.Fields("numnf")
                de_informa.rsSel_GerarCanhotosNF.MoveNext
            Loop
            'final do relatório, fecha com rodapé
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(140, " ")
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.Print Spc(7); "Quantidade de Canhotos da NFS: "; xcontnf
            Printer.Print Spc(7); "Data: "; datahora("data"); Space(35); "Assinatura Conferência: ______________________"
            Printer.NewPage
        Next
        Printer.EndDoc   'finaliza spool da impressão
        MsgBox "Relatório Enviado para Impressão. PROTOCOLO: " & zeros(xnumprot, 6)
    End If
Else

    If de_informa.rsSel_REGerarCanhotosNF.State = 1 Then de_informa.rsSel_REGerarCanhotosNF.Close
    de_informa.Sel_REGerarCanhotosNF Val(txtNumProt)
    If de_informa.rsSel_REGerarCanhotosNF.RecordCount < 1 Then
        MsgBox "Não Há Dados à Serem Impressos para Este Protocolo / Usuário!"
        Exit Sub
    Else
        xnumprot = Val(txtNumProt)
        
        For xcopias = 1 To Val(txtcopias)
            de_informa.rsSel_REGerarCanhotosNF.MoveFirst
            xcontnf = 0
            xColuna = 1
            xlinha = 0
            Do Until de_informa.rsSel_REGerarCanhotosNF.EOF
                xcontnf = xcontnf + 1   'contador de quantidade
                If xlinha = 0 And xColuna = 1 Then   'identifica inicio da página/cabeçário
                    Printer.Print
                    Printer.Print
                    Printer.FontSize = 12
                    Printer.FontBold = True
                    Printer.FontUnderline = True
                    Printer.Print Spc(5); "INTEC TRANSPORTES"
                    Printer.FontUnderline = False
                    Printer.Print
                    Printer.Print Spc(5); "RELATÓRIO DE CANHOTOS DAS NOTAS FISCAIS DEVOLVIDOS PARA O CLIENTE"
                    Printer.Print Spc(5); "DATA: " & de_informa.rsSel_REGerarCanhotosNF.Fields("canhotonfdata")
                    Printer.Print Spc(5); "USUÁRIO: " & de_informa.rsSel_REGerarCanhotosNF.Fields("usu_bx")
                    Printer.Print Spc(5); "PROTOCOLO NÚMERO: " & zeros(xnumprot, 6)
                    Printer.Print Spc(5); "CLIENTE: " & Trim$(de_informa.rsSel_REGerarCanhotosNF.Fields("cliente_nome"))
                    Printer.FontStrikethru = True
                    Printer.Print Spc(5); String(140, " ")
                    Printer.FontSize = 10
                    Printer.FontStrikethru = False
                    Printer.FontBold = False
                    Printer.FontUnderline = False
                End If
                'impressão por 6 colunas
                If xColuna = 1 Then
                    Printer.Print Spc(6); zeros(de_informa.rsSel_REGerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                    xlinha = xlinha + 1
                ElseIf xColuna = 2 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_REGerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 3 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_REGerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 4 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_REGerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 5 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_REGerarCanhotosNF.Fields("numnf"), 10);
                    xColuna = xColuna + 1
                ElseIf xColuna = 6 Then
                    Printer.Print Spc(3); zeros(de_informa.rsSel_REGerarCanhotosNF.Fields("numnf"), 10);
                    'na coluna 6 volta para a coluna 1
                    xColuna = 1
                    Printer.Print
                    'e se a linha for = 26 e não for último CTC ...
                    If xlinha = 26 And de_informa.rsSel_REGerarCanhotosNF.RecordCount <> xcontnf Then
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
                        Printer.Print Spc(7); "Visto Conferência: ______________"
                        Printer.NewPage
                    Else
                        'Printer.Print
                    End If
                End If
                de_informa.rsSel_REGerarCanhotosNF.MoveNext
            Loop

            'final do relatório, fecha com rodapé
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(140, " ")
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.Print Spc(7); "Quantidade de Canhotos da NFS: "; xcontnf
            Printer.Print Spc(7); "Data: "; datahora("data"); Space(35); "Assinatura Conferência: ______________________"
            Printer.NewPage
        Next
        Printer.EndDoc   'finaliza spool da impressão
        MsgBox "Relatório Enviado para Impressão. PROTOCOLO: " & zeros(xnumprot, 6)
    End If








End If

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGeraProtCanhotos = Nothing
End Sub

Private Sub optExcluir_Click()
    txtIncProtocolo.Enabled = False
    txtIncFilial.Enabled = False
    txtIncCtc.Enabled = False
    txtIncCanhoto.Enabled = False
    txtExcProtocolo.Enabled = True
    txtExcCanhoto.Enabled = True
    txtIncProtocolo.BackColor = &H80000005 'branco
    txtIncFilial.BackColor = &H80000005
    txtIncCtc.BackColor = &H80000005
    txtIncCanhoto.BackColor = &H80000005
    txtExcProtocolo.BackColor = &HC0FFFF 'amarelo
    txtExcCanhoto.BackColor = &HC0FFFF
End Sub

Private Sub optGerar_Click()
    txtRemetCGC.Enabled = True
    txtNumProt.Enabled = False
    chkTodosEstab.Enabled = True
    txtRemetCGC.BackColor = &HC0FFFF 'amarelo
    txtNumProt.BackColor = &H80000005 'branco
End Sub

Private Sub optIncluir_Click()
    txtIncProtocolo.Enabled = True
    txtIncFilial.Enabled = True
    txtIncCtc.Enabled = True
    txtIncCanhoto.Enabled = True
    txtExcProtocolo.Enabled = False
    txtExcCanhoto.Enabled = False
    txtIncProtocolo.BackColor = &HC0FFFF 'amarelo
    txtIncFilial.BackColor = &HC0FFFF
    txtIncCtc.BackColor = &HC0FFFF
    txtIncCanhoto.BackColor = &HC0FFFF
    txtExcProtocolo.BackColor = &H80000005 'branco
    txtExcCanhoto.BackColor = &H80000005
End Sub

Private Sub optRegerar_Click()
    txtRemetCGC.Enabled = False
    txtNumProt.Enabled = True
    chkTodosEstab.Enabled = False
    txtRemetCGC.BackColor = &H80000005 'branco
    txtNumProt.BackColor = &HC0FFFF 'amarelo
End Sub

Private Sub txtExcCanhoto_GotFocus()
    txtExcCanhoto.SelStart = 0
    txtExcCanhoto.SelLength = 6
End Sub

Private Sub txtExcCanhoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtExcProtocolo_GotFocus()
    txtExcProtocolo.SelStart = 0
    txtExcProtocolo.SelLength = 6
End Sub

Private Sub txtExcProtocolo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRemetCGC_LostFocus()
    If txtRemetCGC.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(txtRemetCGC) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblRemetNome.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
        Else
            txtRemetCGC.SetFocus
        End If
    Else
        lblRemetNome.Caption = ""
    End If

End Sub
