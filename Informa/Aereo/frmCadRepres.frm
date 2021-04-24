VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadRepres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Representantes"
   ClientHeight    =   6585
   ClientLeft      =   930
   ClientTop       =   1605
   ClientWidth     =   10350
   ControlBox      =   0   'False
   Icon            =   "frmCadRepres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10350
   Begin VB.CommandButton cmd_Canc 
      Caption         =   "Canc/Sair"
      Height          =   435
      Left            =   7740
      TabIndex        =   25
      Top             =   6000
      Width           =   2475
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5220
      TabIndex        =   24
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmd_Alterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   2700
      TabIndex        =   2
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmd_nova 
      Caption         =   "Novo"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   2595
   End
   Begin VB.Frame fra_Dados 
      Caption         =   "Dados do Representante"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   5220
      TabIndex        =   27
      Top             =   60
      Width           =   4995
      Begin VB.TextBox TxtFAX 
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Top             =   3780
         Width           =   1215
      End
      Begin VB.TextBox txtTelRes 
         Height          =   285
         Left            =   2460
         TabIndex        =   16
         Top             =   3780
         Width           =   1155
      End
      Begin VB.TextBox txtTelCel 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   3780
         Width           =   1155
      End
      Begin VB.TextBox txtTelCom 
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   3780
         Width           =   1155
      End
      Begin VB.TextBox TxtCidadeRetira 
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   3240
         Width           =   4155
      End
      Begin VB.TextBox TxtUFRetira 
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   1620
         Width           =   4635
      End
      Begin VB.TextBox txtInscrMun 
         Height          =   285
         Left            =   3300
         TabIndex        =   6
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtInscrEst 
         Height          =   285
         Left            =   1740
         TabIndex        =   5
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtCGC 
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtNomeRepres 
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   540
         Width           =   4635
      End
      Begin VB.TextBox txtAgencia 
         Height          =   285
         Left            =   2100
         TabIndex        =   22
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox txtConta 
         Height          =   285
         Left            =   3480
         TabIndex        =   23
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox txtLocalidade 
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   2700
         Width           =   4155
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   4320
         TabIndex        =   11
         Top             =   2700
         Width           =   495
      End
      Begin VB.TextBox txtNomeBanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         TabIndex        =   20
         Top             =   4860
         Width           =   3795
      End
      Begin VB.TextBox txtNumBanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Top             =   4860
         Width           =   795
      End
      Begin VB.CommandButton cmdBuscaBco 
         Caption         =   "Buscar Banco"
         Height          =   345
         Left            =   180
         TabIndex        =   21
         Top             =   5340
         Width           =   1815
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   4320
         Width           =   4635
      End
      Begin VB.TextBox TxtCEP 
         Height          =   285
         Left            =   3540
         TabIndex        =   9
         Top             =   2160
         Width           =   1275
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   2160
         Width           =   3315
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Cidade onde Retira as Cargas"
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   3000
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   4440
         TabIndex        =   48
         Top             =   3000
         Width           =   210
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nome Banco"
         Height          =   195
         Left            =   1020
         TabIndex        =   47
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Num. Bco."
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   4620
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Telefone Cel."
         Height          =   195
         Left            =   1320
         TabIndex        =   45
         Top             =   3540
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Telefone Res."
         Height          =   195
         Left            =   2460
         TabIndex        =   44
         Top             =   3540
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Telefone Com."
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   3540
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Municipal"
         Height          =   195
         Left            =   3300
         TabIndex        =   41
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Estadual"
         Height          =   195
         Left            =   1740
         TabIndex        =   40
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "CNPJ/CPF"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Nome"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nº Agência"
         Height          =   195
         Left            =   2100
         TabIndex        =   37
         Top             =   5160
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nº Conta"
         Height          =   195
         Left            =   3480
         TabIndex        =   36
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   4440
         TabIndex        =   34
         Top             =   2460
         Width           =   210
      End
      Begin VB.Label lbl_proc 
         Caption         =   "0"
         Height          =   195
         Left            =   7500
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbl_codigo 
         Height          =   255
         Left            =   9360
         TabIndex        =   32
         Top             =   -60
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   4080
         Width           =   435
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         Height          =   195
         Left            =   3540
         TabIndex        =   30
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   3660
         TabIndex        =   28
         Top             =   3540
         Width           =   300
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   5715
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   10081
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label lbl_status 
      Caption         =   "INICIO"
      Height          =   195
      Left            =   8640
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmCadRepres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xSigla As String

Private Sub cmd_adicionar_local_Click()
frmCadLocalidade.Show 1
ComboLocalidade.Clear

If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
    de_informa.Sel_CadLocalAir "%"
        Do Until de_informa.rsSel_CadLocalAir.EOF
        ComboLocalidade.AddItem de_informa.rsSel_CadLocalAir.Fields("localidade") & String(40 - Len(de_informa.rsSel_CadLocalAir.Fields("localidade")), " ") & " - " & de_informa.rsSel_CadLocalAir.Fields("sigla") & " - " & de_informa.rsSel_CadLocalAir.Fields("uf")
        de_informa.rsSel_CadLocalAir.MoveNext
        Loop
End Sub

Private Sub cmd_Alterar_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
Flex.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "ALTERACAO"

txtNomeRepres.BackColor = &HC0FFFF
txtCGC.BackColor = &HC0FFFF
txtInscrEst.BackColor = &HC0FFFF
txtInscrMun.BackColor = &HC0FFFF
txtEndereco.BackColor = &HC0FFFF
txtTelCom.BackColor = &HC0FFFF
txtTelCel.BackColor = &HC0FFFF
txtTelRes.BackColor = &HC0FFFF
txtLocalidade.BackColor = &HC0FFFF
TxtUF.BackColor = &HC0FFFF
TxtCidadeRetira.BackColor = &HC0FFFF
TxtUFRetira.BackColor = &HC0FFFF
txtNumBanco.BackColor = &HC0FFFF
txtNomeBanco.BackColor = &HC0FFFF
txtAgencia.BackColor = &HC0FFFF
txtConta.BackColor = &HC0FFFF
TxtEmail.BackColor = &HC0FFFF
TxtBairro.BackColor = &HC0FFFF
TxtCEP.BackColor = &HC0FFFF
TxtFAX.BackColor = &HC0FFFF
txtNomeRepres.SetFocus
End Sub

Private Sub cmd_Canc_Click()

    If lbl_status.Caption = "INCLUSAO" Or lbl_status.Caption = "ALTERACAO" Then
    cmd_nova.Enabled = True
    Flex.Enabled = True
    fra_Dados.Enabled = False
    cmd_Gravar.Enabled = False
    
    txtNomeRepres.BackColor = &H80000014
    txtCGC.BackColor = &H80000014
    txtInscrEst.BackColor = &H80000014
    txtInscrMun.BackColor = &H80000014
    txtEndereco.BackColor = &H80000014
    txtTelCom.BackColor = &H80000014
    txtTelCel.BackColor = &H80000014
    txtTelRes.BackColor = &H80000014
    txtLocalidade.BackColor = &H80000014
    TxtUF.BackColor = &H80000014
    TxtCidadeRetira.BackColor = &H80000014
    TxtUFRetira.BackColor = &H80000014
    TxtEmail.BackColor = &H80000014
    TxtCEP.BackColor = &H80000014
    TxtBairro.BackColor = &H80000014
    TxtFAX.BackColor = &H80000014
    
    txtNumBanco.BackColor = &H80000014
    txtNomeBanco.BackColor = &H80000014
    txtAgencia.BackColor = &H80000014
    txtConta.BackColor = &H80000014
    
    txtNomeRepres.Text = ""
    txtCGC.Text = ""
    txtInscrEst.Text = ""
    txtInscrMun.Text = ""
    txtEndereco.Text = ""
    txtTelCom.Text = ""
    txtTelCel.Text = ""
    txtTelRes.Text = ""
    txtLocalidade.Text = ""
    TxtUF.Text = ""
    TxtCidadeRetira.Text = ""
    TxtUFRetira.Text = ""
    TxtEmail.Text = ""
    TxtCEP.Text = ""
    TxtBairro.Text = ""
    TxtFAX.Text = ""
    txtNumBanco.Text = ""
    txtNomeBanco.Text = ""
    txtAgencia.Text = ""
    txtConta.Text = ""
    
    
    
    lbl_codigo.Caption = ""
    lbl_status.Caption = "INICIO"
    
    Else
    Unload Me
    End If
End Sub



Private Sub cmd_Gravar_Click()
Dim xControl As Control
Dim xControlText As String

    If Trim(txtNomeRepres.Text) = "" Then
    Set xControl = txtNomeRepres
    xControlText = "Nome do Representante"
    MsgBox "O Campo " & xControlText & " não foi preenchido corretamente. Por favor tente novamente...", vbExclamation, ""
    xControl.SetFocus
    Exit Sub
    ElseIf Trim(txtCGC.Text) = "" Then
    Set xControl = txtCGC
    xControlText = "CNPJ/CPF"
    MsgBox "O Campo " & xControlText & " não foi preenchido corretamente. Por favor tente novamente...", vbExclamation, ""
    xControl.SetFocus
    Exit Sub
    ElseIf Trim(txtTelCom.Text) = "" And Trim(txtTelCel.Text) = "" And Trim(txtTelRes.Text) = "" Then
    MsgBox "Nenhum Campo de Telefone foi preenchido. Ao menos um Telefone para contato é necessário. Por favor tente novamente...", vbExclamation, ""
    Exit Sub
    ElseIf Trim(txtLocalidade.Text) = "" Then
    Set xControl = txtLocalidade
    xControlText = "Cidade"
    MsgBox "O Campo " & xControlText & " não foi preenchido corretamente. Por favor tente novamente...", vbExclamation, ""
    xControl.SetFocus
    Exit Sub
    ElseIf Trim(TxtUF.Text) = "" Then
    Set xControl = TxtUF
    xControlText = "UF"
    MsgBox "O Campo " & xControlText & " não foi preenchido corretamente. Por favor tente novamente...", vbExclamation, ""
    xControl.SetFocus
    Exit Sub
    End If

    If Len(Trim(TxtCidadeRetira.Text)) = 0 Then TxtCidadeRetira.Text = txtLocalidade.Text
    If Len(Trim(TxtUFRetira.Text)) = 0 Then TxtUFRetira.Text = TxtUF.Text
    
    If de_informa.rsSel_CONFCidade.State = 1 Then de_informa.rsSel_CONFCidade.Close
    de_informa.Sel_CONFCidade Trim(txtLocalidade.Text) & "%"
    
    If de_informa.rsSel_CONFCidade.RecordCount = 0 Then
    MsgBox "Nenhuma correpondência de Cidade foi encontrada para a digitada. Por favor reveja este Campo.", vbInformation, ""
    txtLocalidade.SetFocus
    Exit Sub
    'ElseIf de_informa.rsSel_CONFCidade.RecordCount > 1 Then
    '    If MsgBox("Foram encontradas " & de_informa.rsSel_CONFCidade.RecordCount & " referências semelhantes à esta Cidade que você digitou. Clique OK para escolher a Cidade correta.", vbOKCancel + vbExclamation, "") = vbOK Then
    '    Set xForm = Me
    '    FrmCadLocalidadeCONF.Show 1
    '    Else
    '    txtLocalidade.SetFocus
    '    Exit Sub
    '    End If
    End If
    
    

cmd_Gravar.Enabled = False
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False

    If lbl_status.Caption = "INCLUSAO" Then
    If de_informa.rsSel_CadRepresCodigo.State = 1 Then de_informa.rsSel_CadRepresCodigo.Close
    de_informa.Sel_CadRepresCodigo
    
    lbl_codigo.Caption = Trim(Str(de_informa.rsSel_CadRepresCodigo.Fields("codigo") + 1))
    de_informa.Ins_CadRepres Val(lbl_codigo.Caption), UCase(Trim(txtNomeRepres.Text)), Trim(txtCGC.Text), Trim(txtInscrEst.Text), Trim(txtInscrMun), UCase(Trim(txtLocalidade.Text)), UCase(Trim(TxtUF.Text)), UCase(Trim(txtEndereco.Text)), LCase(Trim(TxtEmail.Text)), UCase(Trim(TxtBairro.Text)), Trim(TxtCEP.Text), Trim(txtTelCom.Text), Trim(txtTelCel.Text), Trim(txtTelRes.Text), Trim(TxtFAX.Text), UCase(Trim(txtNomeBanco.Text)), Trim(txtNumBanco.Text), Trim(txtAgencia.Text), Trim(txtConta.Text), Trim(UCase(TxtCidadeRetira.Text)), Trim(UCase(TxtUFRetira.Text))
    lbl_codigo.Caption = ""
    
    ElseIf lbl_status.Caption = "ALTERACAO" Then
    de_informa.Update_CadRepres UCase(Trim(txtNomeRepres.Text)), Trim(txtCGC.Text), Trim(txtInscrEst.Text), Trim(txtInscrMun), UCase(Trim(txtLocalidade.Text)), UCase(Trim(TxtUF.Text)), UCase(Trim(txtEndereco.Text)), LCase(Trim(TxtEmail.Text)), UCase(Trim(TxtBairro.Text)), Trim(TxtCEP.Text), Trim(txtTelCom.Text), Trim(txtTelCel.Text), Trim(txtTelRes.Text), Trim(TxtFAX.Text), UCase(Trim(txtNomeBanco.Text)), Trim(txtNumBanco.Text), Trim(txtAgencia.Text), Trim(txtConta.Text), Trim(UCase(TxtCidadeRetira.Text)), Trim(UCase(TxtUFRetira.Text)), Val(lbl_codigo.Caption)
    lbl_codigo.Caption = ""
    End If
    
cmd_nova.Enabled = True
Flex.Enabled = True
fra_Dados.Enabled = False
cmd_Gravar.Enabled = False



txtNomeRepres.BackColor = &H80000014
txtCGC.BackColor = &H80000014
txtInscrEst.BackColor = &H80000014
txtInscrMun.BackColor = &H80000014
txtEndereco.BackColor = &H80000014
txtTelCom.BackColor = &H80000014
txtTelCel.BackColor = &H80000014
txtTelRes.BackColor = &H80000014
txtLocalidade.BackColor = &H80000014
TxtUF.BackColor = &H80000014
TxtCidadeRetira.BackColor = &H80000014
TxtUFRetira.BackColor = &H80000014
txtNumBanco.BackColor = &H80000014
txtNomeBanco.BackColor = &H80000014
txtAgencia.BackColor = &H80000014
txtConta.BackColor = &H80000014
TxtEmail.BackColor = &H80000014
TxtCEP.BackColor = &H80000014
TxtBairro.BackColor = &H80000014
TxtFAX.BackColor = &H80000014


txtNomeRepres.Text = ""
txtCGC.Text = ""
txtInscrEst.Text = ""
txtInscrMun.Text = ""
txtEndereco.Text = ""
txtTelCom.Text = ""
txtTelCel.Text = ""
txtTelRes.Text = ""
txtLocalidade.Text = ""
TxtUF.Text = ""
TxtCidadeRetira.Text = ""
TxtUFRetira.Text = ""

txtNumBanco.Text = ""
txtNomeBanco.Text = ""
txtAgencia.Text = ""
txtConta.Text = ""
TxtEmail.Text = ""
TxtCEP.Text = ""
TxtBairro.Text = ""
TxtFAX.Text = ""


Dim x As Integer
        

Call Form_Load

lbl_status.Caption = "INICIO"
End Sub

Private Sub cmd_Nova_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
Flex.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "INCLUSAO"

txtNomeRepres.BackColor = &HC0FFFF
txtCGC.BackColor = &HC0FFFF
txtInscrEst.BackColor = &HC0FFFF
txtInscrMun.BackColor = &HC0FFFF
txtEndereco.BackColor = &HC0FFFF
txtTelCom.BackColor = &HC0FFFF
txtTelCel.BackColor = &HC0FFFF
txtTelRes.BackColor = &HC0FFFF
txtLocalidade.BackColor = &HC0FFFF
TxtUF.BackColor = &HC0FFFF
TxtCidadeRetira.BackColor = &HC0FFFF
TxtUFRetira.BackColor = &HC0FFFF
txtNumBanco.BackColor = &HC0FFFF
txtNomeBanco.BackColor = &HC0FFFF
txtAgencia.BackColor = &HC0FFFF
txtConta.BackColor = &HC0FFFF
TxtEmail.BackColor = &HC0FFFF
TxtBairro.BackColor = &HC0FFFF
TxtCEP.BackColor = &HC0FFFF
TxtFAX.BackColor = &HC0FFFF


txtNomeRepres.Text = ""
txtCGC.Text = ""
txtInscrEst.Text = ""
txtInscrMun.Text = ""
txtEndereco.Text = ""
txtTelCom.Text = ""
txtTelCel.Text = ""
txtTelRes.Text = ""
txtLocalidade.Text = ""
TxtUF.Text = ""
TxtCidadeRetira.Text = ""
TxtUFRetira.Text = ""
txtNumBanco.Text = ""
txtNomeBanco.Text = ""
txtAgencia.Text = ""
txtConta.Text = ""
TxtEmail.Text = ""
TxtBairro.Text = ""
TxtCEP.Text = ""
TxtFAX.Text = ""


txtNomeRepres.SetFocus
End Sub





Private Sub ComboBanco_Click()
ComboNumBanco.Text = ComboNumBanco.List(ComboBanco.ListIndex)
End Sub

Private Sub ComboBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    lbl_proc.Caption = "1"
    ElseIf KeyAscii = 13 Then
    lbl_proc.Caption = "0"
    SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdBuscaBco_Click()
Set xForm = Me
txtNomeBanco.Text = ""
txtNumBanco.Text = ""
frmProcBancos.Show 1
Set xForm = Nothing
End Sub

Private Sub ComboLocalidade_Change()
Dim Cont As Integer
Dim xSelStart As Integer

If lbl_proc.Caption = "0" Then
    If Len(Trim(ComboLocalidade.Text)) > 0 Then
        For Cont = 0 To ComboLocalidade.ListCount - 1
        If UCase(ComboLocalidade.Text) = UCase(Mid(ComboLocalidade.List(Cont), 1, Len(ComboLocalidade.Text))) Then
        lbl_proc.Caption = "1"
        xSelStart = Len(ComboLocalidade.Text)
        ComboLocalidade.Text = ComboLocalidade.List(Cont)
        lbl_proc.Caption = "0"
        ComboLocalidade.SelStart = xSelStart
        ComboLocalidade.SelLength = 300
        Cont = ComboLocalidade.ListCount - 1
        End If
        Next
    End If
End If
lbl_proc.Caption = "0"
End Sub

Private Sub ComboLocalidade_GotFocus()
ComboLocalidade.SelStart = 0
ComboLocalidade.SelLength = 500
End Sub

Private Sub ComboLocalidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    lbl_proc.Caption = "1"
    ElseIf KeyAscii = 13 Then
    lbl_proc.Caption = "0"
    SendKeys "{TAB}"
    End If
End Sub

Private Sub ComboNumBanco_Change()
Dim Cont As Integer
Dim xSelStart As Integer

If lbl_proc.Caption = "0" Then
    If Len(Trim(ComboNumBanco.Text)) > 0 Then
        For Cont = 0 To ComboNumBanco.ListCount - 1
        If UCase(ComboNumBanco.Text) = UCase(Mid(ComboNumBanco.List(Cont), 1, Len(ComboNumBanco.Text))) Then
        lbl_proc.Caption = "1"
        xSelStart = Len(ComboNumBanco.Text)
        ComboNumBanco.Text = ComboNumBanco.List(Cont)
        ComboBanco.Text = ComboBanco.List(Cont)
        lbl_proc.Caption = "0"
        ComboNumBanco.SelStart = xSelStart
        ComboNumBanco.SelLength = 300
        Cont = ComboNumBanco.ListCount - 1
        Else
        lbl_proc.Caption = "1"
        ComboBanco.Text = ""
        lbl_proc.Caption = "0"
        End If
        Next
    End If
Else
ComboBanco.Text = ""
End If
lbl_proc.Caption = "0"
End Sub

Private Sub ComboNumBanco_Click()
ComboBanco.Text = ComboBanco.List(ComboNumBanco.ListIndex)
End Sub

Private Sub ComboNumBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
    lbl_proc.Caption = "1"
    ElseIf KeyAscii = 13 Then
    lbl_proc.Caption = "0"
    SendKeys "{TAB}"
    End If
End Sub



Private Sub Flex_Click()
txtLocalidade.Text = Flex.TextMatrix(Flex.Row, 0)
TxtUF.Text = Flex.TextMatrix(Flex.Row, 1)
txtNomeRepres.Text = Flex.TextMatrix(Flex.Row, 2)
txtCGC.Text = Flex.TextMatrix(Flex.Row, 3)
txtInscrEst.Text = Flex.TextMatrix(Flex.Row, 4)
txtInscrMun.Text = Flex.TextMatrix(Flex.Row, 5)
txtEndereco.Text = Flex.TextMatrix(Flex.Row, 6)
TxtBairro.Text = Flex.TextMatrix(Flex.Row, 7)
TxtCEP.Text = Flex.TextMatrix(Flex.Row, 8)
txtTelCom.Text = Flex.TextMatrix(Flex.Row, 9)
txtTelCel.Text = Flex.TextMatrix(Flex.Row, 10)
txtTelRes.Text = Flex.TextMatrix(Flex.Row, 11)
TxtFAX.Text = Flex.TextMatrix(Flex.Row, 12)
TxtEmail.Text = Flex.TextMatrix(Flex.Row, 13)
txtNumBanco.Text = Flex.TextMatrix(Flex.Row, 14)
txtNomeBanco.Text = Flex.TextMatrix(Flex.Row, 15)
txtAgencia.Text = Flex.TextMatrix(Flex.Row, 16)
txtConta.Text = Flex.TextMatrix(Flex.Row, 17)
lbl_codigo.Caption = Flex.TextMatrix(Flex.Row, 18)
TxtCidadeRetira.Text = Flex.TextMatrix(Flex.Row, 19)
TxtUFRetira.Text = Flex.TextMatrix(Flex.Row, 20)

cmd_Alterar.Enabled = True
End Sub

Private Sub Flex_SelChange()
Flex_Click
End Sub

Private Sub Form_Load()
Dim x As Integer
If de_informa.rsSel_CadRepres.State = 1 Then de_informa.rsSel_CadRepres.Close
de_informa.Sel_CadRepres

Flex.Clear
Flex.Rows = de_informa.rsSel_CadRepres.RecordCount + 1
Flex.Cols = 21
Flex.FixedCols = 0
Flex.FixedRows = 1
Flex.TextMatrix(0, 0) = "Cidade"
Flex.TextMatrix(0, 1) = "UF"
Flex.TextMatrix(0, 2) = "Nome"
Flex.TextMatrix(0, 3) = "CNPJ/CPF"
Flex.TextMatrix(0, 4) = "Ins. Est."
Flex.TextMatrix(0, 5) = "Ins. Mun."
Flex.TextMatrix(0, 6) = "Endereço"
Flex.TextMatrix(0, 7) = "Bairro"
Flex.TextMatrix(0, 8) = "CEP"
Flex.TextMatrix(0, 9) = "Tel. Com."
Flex.TextMatrix(0, 10) = "Tel. Cel."
Flex.TextMatrix(0, 11) = "Tel. Res."
Flex.TextMatrix(0, 12) = "Fax"
Flex.TextMatrix(0, 13) = "E-Mail"
Flex.TextMatrix(0, 14) = "Num. Bco."
Flex.TextMatrix(0, 15) = "Nome Bco."
Flex.TextMatrix(0, 16) = "Ag."
Flex.TextMatrix(0, 17) = "C/C"
Flex.TextMatrix(0, 18) = "Cód."
Flex.TextMatrix(0, 19) = "Cidade Retira"
Flex.TextMatrix(0, 20) = "UF Retira"


Flex.ColWidth(0) = 3000
Flex.ColWidth(1) = 500
Flex.ColWidth(2) = 5000
Flex.ColWidth(3) = 1500
Flex.ColWidth(4) = 1500
Flex.ColWidth(5) = 1500
Flex.ColWidth(6) = 5000
Flex.ColWidth(7) = 1500
Flex.ColWidth(8) = 1000
Flex.ColWidth(9) = 1000
Flex.ColWidth(10) = 1000
Flex.ColWidth(11) = 1000
Flex.ColWidth(12) = 1000
Flex.ColWidth(13) = 1500
Flex.ColWidth(14) = 1000
Flex.ColWidth(15) = 3000
Flex.ColWidth(16) = 1500
Flex.ColWidth(17) = 1500
Flex.ColWidth(18) = 900
Flex.ColWidth(19) = 3000
Flex.ColWidth(20) = 500


x = 1
With de_informa.rsSel_CadRepres
    Do Until .EOF
    If IsNull(.Fields("localidade")) = False Then Flex.TextMatrix(x, 0) = PriMaiuscula(.Fields("localidade"))
    If IsNull(.Fields("uf")) = False Then Flex.TextMatrix(x, 1) = .Fields("uf")
    If IsNull(.Fields("nome")) = False Then Flex.TextMatrix(x, 2) = PriMaiuscula(.Fields("nome"))
    If IsNull(.Fields("cgc")) = False Then Flex.TextMatrix(x, 3) = .Fields("cgc")
    If IsNull(.Fields("inscr_est")) = False Then Flex.TextMatrix(x, 4) = .Fields("inscr_est")
    If IsNull(.Fields("inscr_mun")) = False Then Flex.TextMatrix(x, 5) = .Fields("inscr_mun")
    If IsNull(.Fields("endereco")) = False Then Flex.TextMatrix(x, 6) = PriMaiuscula(.Fields("endereco"))
    If IsNull(.Fields("bairro")) = False Then Flex.TextMatrix(x, 7) = PriMaiuscula(.Fields("bairro"))
    If IsNull(.Fields("cep")) = False Then Flex.TextMatrix(x, 8) = .Fields("cep")
    If IsNull(.Fields("telcom")) = False Then Flex.TextMatrix(x, 9) = .Fields("telcom")
    If IsNull(.Fields("telcel")) = False Then Flex.TextMatrix(x, 10) = .Fields("telcel")
    If IsNull(.Fields("telres")) = False Then Flex.TextMatrix(x, 11) = .Fields("telres")
    If IsNull(.Fields("fax")) = False Then Flex.TextMatrix(x, 12) = .Fields("fax")
    If IsNull(.Fields("email")) = False Then Flex.TextMatrix(x, 13) = .Fields("email")
    If IsNull(.Fields("banconum")) = False Then Flex.TextMatrix(x, 14) = .Fields("banconum")
    If IsNull(.Fields("banco")) = False Then Flex.TextMatrix(x, 15) = PriMaiuscula(.Fields("banco"))
    If IsNull(.Fields("agencia")) = False Then Flex.TextMatrix(x, 16) = .Fields("agencia")
    If IsNull(.Fields("conta")) = False Then Flex.TextMatrix(x, 17) = .Fields("conta")
    If IsNull(.Fields("codigo")) = False Then Flex.TextMatrix(x, 18) = .Fields("codigo")
    If IsNull(.Fields("cidaderetira")) = False Then Flex.TextMatrix(x, 19) = PriMaiuscula(.Fields("cidaderetira"))
    If IsNull(.Fields("ufretira")) = False Then Flex.TextMatrix(x, 20) = .Fields("ufretira")
    x = x + 1
    .MoveNext
    Loop
End With
        

        
End Sub

Private Sub Teste_Change()

Dim Abc As String
Dim Cont As Integer
Dim Flag As Boolean
Dim TextoAux As String

Abc = "ABCEDFGHIJKLMNOPQRSTUVYXWZ"
TextoAux = ""

    If Len(TxtSigla.Text) > 0 Then
        
        Flag = False
        
        For Cont = 1 To Len(Abc)
            If UCase(Mid(TxtSigla.Text, TxtSigla.SelStart, 1)) = Mid(Abc, Cont, 1) Then
            Flag = True
            End If
        Next
    
        If Flag = False Then
            For Cont = 1 To Len(TxtSigla.Text)
                If Cont <> TxtSigla.SelStart Then
                TextoAux = Mid(TxtSigla, Cont, 1)
                End If
            Next
        Else
        TextoAux = TxtSigla.Text
        End If
        
        TxtSigla.Text = TextoAux
    End If
        
End Sub

Private Sub GridCadRepres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GridCadrepres_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If lbl_status.Caption = "INICIO" Then
    txtNomeRepres.Text = GridCadRepres.Columns(3)
    txtCGC.Text = GridCadRepres.Columns(4)
    txtInscrEst.Text = GridCadRepres.Columns(5)
    txtInscrMun.Text = GridCadRepres.Columns(6)
    txtEndereco.Text = GridCadRepres.Columns(7)
    txtTelCom.Text = GridCadRepres.Columns(8)
    txtTelCel.Text = GridCadRepres.Columns(9)
    txtTelRes.Text = GridCadRepres.Columns(10)
    txtLocalidade.Text = GridCadRepres.Columns(1)
    TxtUF.Text = GridCadRepres.Columns(2)
    
    txtNumBanco.Text = GridCadRepres.Columns(12)
    txtNomeBanco.Text = GridCadRepres.Columns(11)
    txtAgencia.Text = GridCadRepres.Columns(13)
    txtConta.Text = GridCadRepres.Columns(14)
    lbl_codigo.Caption = GridCadRepres.Columns(15)
    cmd_Alterar.Enabled = True
    End If

End Sub


Private Sub TxtCidadeRetira_GotFocus()
TxtCidadeRetira.SelStart = 0
TxtCidadeRetira.SelLength = 150
End Sub

Private Sub Txtlocalidade_GotFocus()
txtLocalidade.SelStart = 0
txtLocalidade.SelLength = 100
End Sub

Private Sub Txtlocalidade_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub txtLocalidade_LostFocus()
txtLocalidade.Text = UCase(txtLocalidade.Text)
End Sub

Private Sub TxtSigla_GotFocus()
TxtSigla.SelStart = 0
TxtSigla.SelLength = 10
End Sub

Private Sub TxtSigla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSigla_LostFocus()
TxtSigla.Text = UCase(TxtSigla.Text)
End Sub

Private Sub GridCadRepres_Click()
txtLocalidade.Text = GridCadRepres.Columns(0)
TxtUF.Text = GridCadRepres.Columns(0)
txtNomeRepres.Text = GridCadRepres.Columns(0)
txtCGC.Text = GridCadRepres.Columns(0)
txtInscrEst.Text = GridCadRepres.Columns(0)
txtInscrMun.Text = GridCadRepres.Columns(0)
txtEndereco.Text = GridCadRepres.Columns(0)
TxtEmail.Text = GridCadRepres.Columns(0)
TxtBairro.Text = GridCadRepres.Columns(0)
TxtCEP.Text = GridCadRepres.Columns(0)
txtTelCom.Text = GridCadRepres.Columns(0)
txtTelCel.Text = GridCadRepres.Columns(0)
txtTelRes.Text = GridCadRepres.Columns(0)
TxtFAX.Text = GridCadRepres.Columns(0)
txtNumBanco.Text = GridCadRepres.Columns(0)
txtNomeBanco.Text = GridCadRepres.Columns(0)
txtAgencia.Text = GridCadRepres.Columns(0)
txtConta.Text = GridCadRepres.Columns(0)
lbl_codigo.Caption = GridCadRepres.Columns(0)
cmd_Alterar.Enabled = True
End Sub

Private Sub txtAgencia_GotFocus()
txtAgencia.SelStart = 0
txtAgencia.SelLength = 500
End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtCEP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtCGC_GotFocus()
txtCGC.SelStart = 0
txtCGC.SelLength = 500
End Sub

Private Sub txtCGC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

'Private Sub Txtlocalidade_GotFocus()
'txtLocalidade.SelStart = 0
'txtLocalidade.SelLength = 500
'End Sub

'Private Sub Txtlocalidade_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'SendKeys "{TAB}"
'KeyAscii = 0
'End If
'End Sub

Private Sub txtConta_GotFocus()
txtConta.SelStart = 0
txtConta.SelLength = 500
End Sub

Private Sub txtConta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtEndereco_GotFocus()
txtEndereco.SelStart = 0
txtEndereco.SelLength = 500
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtFAX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtInscrEst_GotFocus()
txtInscrEst.SelStart = 0
txtInscrEst.SelLength = 500
End Sub

Private Sub txtInscrEst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtInscrMun_GotFocus()
txtInscrMun.SelStart = 0
txtInscrMun.SelLength = 500
End Sub

Private Sub txtInscrMun_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtNomeRepres_GotFocus()
txtNomeRepres.SelStart = 0
txtNomeRepres.SelLength = 500
End Sub

Private Sub txtNomeRepres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtNumBanco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtTelCel_GotFocus()
txtTelCel.SelStart = 0
txtTelCel.SelLength = 500
End Sub

Private Sub txtTelCel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtTelCom_GotFocus()
txtTelCom.SelStart = 0
txtTelCom.SelLength = 500
End Sub

Private Sub txtTelCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtTelRes_GotFocus()
txtTelRes.SelStart = 0
txtTelRes.SelLength = 500
End Sub

Private Sub txtTelRes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtUF_GotFocus()
TxtUF.SelStart = 0
TxtUF.SelLength = 10
End Sub

Private Sub txtUF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtUF_LostFocus()
TxtUF.Text = UCase(TxtUF.Text)
End Sub

Private Sub TxtUFRetira_GotFocus()
TxtUFRetira.SelStart = 0
TxtUFRetira.SelLength = 150
End Sub
