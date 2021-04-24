VERSION 5.00
Begin VB.Form frmCancCTC 
   Caption         =   "Cancelar CTC"
   ClientHeight    =   4830
   ClientLeft      =   2250
   ClientTop       =   1440
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6255
   Begin VB.Frame Frame2 
      Caption         =   "CTCs Cancelados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdDescancelar 
         Caption         =   "Retirar Cancelamento"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "Procurar..."
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair2 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Confirma Cancelar CTC"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox txtObsCanc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   3
         Top             =   3480
         Width           =   5775
      End
      Begin VB.TextBox txtCTC 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtfilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fatura:"
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblFatura 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblEntregueSN 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   24
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblCidadeDest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lblUFdest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5040
         TabIndex        =   21
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   4680
         TabIndex        =   20
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cidade Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label lblPeso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblValNf 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Left            =   3600
         TabIndex        =   16
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valor NF:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblDestinataio 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label lblRemetente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Destinatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observação de Cancelamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   2205
      End
      Begin VB.Label Label3 
         Caption         =   "Número da Filial/CTC:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCancCTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    If Len(Trim$(txtObsCanc)) < 3 Then
        MsgBox "Coloque o Motivo do Cancelamento !", vbCritical, "Observação"
        txtObsCanc.SetFocus
    Else
        If MsgBox("Confirma Cancelar este CTC ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            de_informa.Alt_CancCTC xusuario, txtObsCanc, transctc(TxtFilial, txtCtc)
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "CANCELAMENTO DE CTC: " & transctc(TxtFilial, txtCtc)
            
            MsgBox "OK ! CTC Cancelado."
            
            cmdProcurar_Click
        End If
    End If
End Sub

Private Sub cmdDescancelar_Click()
    If Mid$(lblEntregueSN, 1, 13) = "CTC CANCELADO" Then
        If MsgBox("Confirma a Retirada de Cancelamento deste CTC ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            de_informa.Alt_Descancelar zeros(TxtFilial, 2) & zeros(txtCtc, 8)
        
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "RETIRADA DE CANCELAMENTO. CTC: " & transctc(TxtFilial, txtCtc)
        
            MsgBox "OK ! Retirado Cancelamento deste CTC."
            
            cmdProcurar_Click
        End If
    Else
        MsgBox "ERRO ! Este CTC não está Cancelado !!!"
    End If
End Sub

Private Sub cmdProcurar_Click()
    lblEmissao = ""
    lblRemetente = ""
    lblDestinataio = ""
    lblCidadeDest = ""
    lblUfDest = ""
    lblValNf = ""
    lblPeso = ""
    lblEntregueSN = ""
    lblFatura = ""
    CmdCancelar.Enabled = False
    cmdDescancelar.Enabled = False
    txtObsCanc.Enabled = False
    txtObsCanc.BackColor = &H8000000E  'branco
    
    If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
    de_informa.Sel_Ctc_SAC transctc(TxtFilial, txtCtc)
    
    If de_informa.rsSel_Ctc_SAC.EOF Then
        MsgBox "Filial/CTC não Encontrado !"
        TxtFilial.SetFocus
    Else
        lblEmissao = de_informa.rsSel_Ctc_SAC.Fields("data")
        lblRemetente = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
        lblDestinataio = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
        lblCidadeDest = de_informa.rsSel_Ctc_SAC.Fields("cidade_dest")
        lblUfDest = de_informa.rsSel_Ctc_SAC.Fields("uf_dest")
        lblValNf = Format(de_informa.rsSel_Ctc_SAC.Fields("valmerc"), "##,###,##0.00")
        lblPeso = Format(de_informa.rsSel_Ctc_SAC.Fields("peso"), "##,##0.0")
        lblFatura = de_informa.rsSel_Ctc_SAC.Fields("faturanum")
        xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
        lblEntregueSN.ToolTipText = ""
        
        If xtemocorr = "0" Then
            lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
            lblEntregueSN.Caption = "OCORR/Baixado"
            TxtFilial.SetFocus
        ElseIf xtemocorr = "1" Then
            lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
            lblEntregueSN.Caption = "OK. ENTREGUE"
            TxtFilial.SetFocus
        ElseIf xtemocorr = "2" Then
            lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "OCORR/Pendente"
            TxtFilial.SetFocus
        ElseIf xtemocorr = "N" Then
            CmdCancelar.Enabled = True
            txtObsCanc.Enabled = True
            txtObsCanc.BackColor = &HC0FFFF
            txtObsCanc.SetFocus
            If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= datahora("data") Then
                lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
                lblEntregueSN.Caption = "EM TRÂNSITO"
                lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
            Else
                lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
                lblEntregueSN.Caption = "SEM POSIÇÃO"
                lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
            End If
        ElseIf xtemocorr = "C" Then
            'cmdDescancelar.Enabled = True
            TxtFilial.SetFocus
            lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "CTC CANCELADO"
            lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                        "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                        "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
        End If
        
        If Len(Trim$(lblFatura)) > 0 Then
            CmdCancelar.Enabled = False
        End If
        
        If Mid$(de_informa.rsSel_Ctc_SAC.Fields("respons_cgc"), 1, 8) = "33200056" Then
            lblEntregueSN.Caption = lblEntregueSN.Caption & " (RIACHUELO)"
            CmdCancelar.Enabled = True
            txtObsCanc.Enabled = True
            txtObsCanc.BackColor = &HC0FFFF
            txtObsCanc.SetFocus
        End If
    End If
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub cmdSair2_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCancCTC = Nothing
End Sub

Private Sub txtCtc_Change()
    If Len(txtCtc.Text) >= 8 Then cmdProcurar.SetFocus
End Sub
Private Sub txtCTC_GotFocus()
    txtCtc.SelStart = 0
    txtCtc.SelLength = 8
End Sub
Private Sub txtCTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub TxtFilial_Change()
    If Len(TxtFilial.Text) >= 2 Then txtCtc.SetFocus
End Sub
Private Sub TxtFilial_GotFocus()
    TxtFilial.SelStart = 0
    TxtFilial.SelLength = 2
End Sub
Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtObsCanc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        If CmdCancelar.Enabled = True Then
            SendKeys "{TAB}"  'ENVIA UM TAB
        Else
            SendKeys "{TAB}"  'ENVIA UM TAB
            SendKeys "{TAB}"  'ENVIA +UM TAB
        End If
    End If
End Sub

Private Sub txtSenha_Change()
    txtSenha.SelStart = 0
    txtSenha.SelLength = 10
End Sub
Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtUsuario_Change()
    TxtUsuario.SelStart = 0
    TxtUsuario.SelLength = 10
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
