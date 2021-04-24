VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadProdIATA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Categoria de Produtos - IATA"
   ClientHeight    =   3855
   ClientLeft      =   3030
   ClientTop       =   2565
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmCadProdANVISA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5895
   Begin MSDataGridLib.DataGrid GridCadIATA 
      Bindings        =   "frmCadProdANVISA.frx":000C
      Height          =   1875
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3307
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Sel_CadIATA"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descricao"
         Caption         =   "Descricao"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004,788
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_Alterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1620
      TabIndex        =   4
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3120
      TabIndex        =   6
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Canc 
      Caption         =   "Canc/Sair"
      Height          =   435
      Left            =   4620
      TabIndex        =   7
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Nova 
      Caption         =   "Nova"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
   Begin VB.Frame fra_Dados 
      Caption         =   "Categoria de Produtos"
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
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5655
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   540
         Width           =   4155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição da Categoria"
         Height          =   195
         Left            =   1320
         TabIndex        =   10
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Label lbl_status 
      Caption         =   "INICIO"
      Height          =   195
      Left            =   4740
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmCadProdIATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xCodigo As String
Private Sub cmd_Alterar_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
GridCadIATA.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "ALTERACAO"

xCodigo = GridCadIATA.Columns(0)

txtCodigo.Enabled = False
txtDescricao.BackColor = &HC0FFFF
txtDescricao.SetFocus
End Sub

Private Sub cmd_Canc_Click()

    If lbl_status.Caption = "INCLUSAO" Or lbl_status.Caption = "ALTERACAO" Then
    cmd_nova.Enabled = True
    GridCadIATA.Enabled = True
    fra_Dados.Enabled = False
    cmd_Gravar.Enabled = False
    
    txtDescricao.BackColor = &H80000014
    txtCodigo.BackColor = &H80000014
   
    If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
    de_informa.Sel_Cadiata "%"
    GridCadIATA.DataMember = "Sel_Cadiata"
    GridCadIATA.Refresh
    DoEvents
    
    txtDescricao.Text = ""
    txtCodigo.Text = ""

    lbl_status.Caption = "INICIO"
    
    Else
    Unload Me
    End If
End Sub



Private Sub cmd_Gravar_Click()
    If Len(Trim(txtDescricao.Text)) = 0 Then
    MsgBox "Não foi informada corretamente a Descricao", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(txtCodigo.Text)) = 0 Then
    MsgBox "Não foi informado corretamente o Código", vbCritical, ""
    Exit Sub
    End If

cmd_Gravar.Enabled = False
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False

    If lbl_status.Caption = "INCLUSAO" Then
        If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
        de_informa.Sel_Cadiata Trim(txtCodigo.Text)
        
        If de_informa.rsSel_CadIATA.RecordCount > 0 Then
        MsgBox "Já existe uma Categoria de Produtos cadastrada com este Código. O Casdastramento não será efetuado.", vbExclamation, ""
        Else
        de_informa.Ins_Cadiata Trim(txtCodigo.Text), UCase(Trim(txtDescricao.Text))
        End If
    ElseIf lbl_status.Caption = "ALTERACAO" Then
        
        de_informa.Update_Cadiata Trim(txtCodigo.Text), UCase(Trim(txtDescricao.Text)), xCodigo
       
    End If
    
cmd_nova.Enabled = True
GridCadIATA.Enabled = True
fra_Dados.Enabled = False
cmd_Gravar.Enabled = False

txtDescricao.BackColor = &H80000014
txtDescricao.Text = ""

txtCodigo.BackColor = &H80000014
txtCodigo.Text = ""
txtCodigo.Enabled = True

If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
de_informa.Sel_Cadiata "%"
GridCadIATA.DataMember = "Sel_Cadiata"
GridCadIATA.Refresh
DoEvents

lbl_status.Caption = "INICIO"
End Sub

Private Sub cmd_Nova_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
GridCadIATA.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "INCLUSAO"

txtCodigo.BackColor = &HC0FFFF
txtCodigo.Text = ""

txtDescricao.BackColor = &HC0FFFF
txtDescricao.Text = ""
txtCodigo.SetFocus
End Sub


Private Sub Form_Load()
    
    If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
    de_informa.Sel_Cadiata "%"
    GridCadIATA.DataMember = "Sel_Cadiata"
    GridCadIATA.Refresh
    DoEvents
    
End Sub

Private Sub Teste_Change()

Dim Abc As String
Dim Cont As Integer
Dim Flag As Boolean
Dim TextoAux As String

Abc = "ABCEDFGHIJKLMNOPQRSTUVYXWZ"
TextoAux = ""

    If Len(txtSigla.Text) > 0 Then
        
        Flag = False
        
        For Cont = 1 To Len(Abc)
            If UCase(Mid(txtSigla.Text, txtSigla.SelStart, 1)) = Mid(Abc, Cont, 1) Then
            Flag = True
            End If
        Next
    
        If Flag = False Then
            For Cont = 1 To Len(txtSigla.Text)
                If Cont <> txtSigla.SelStart Then
                TextoAux = Mid(txtSigla, Cont, 1)
                End If
            Next
        Else
        TextoAux = txtSigla.Text
        End If
        
        txtSigla.Text = TextoAux
    End If
        
End Sub

Private Sub GridCadiata_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If lbl_status.Caption = "INICIO" Then
    txtCodigo.Text = GridCadIATA.Columns(0)
    txtDescricao.Text = GridCadIATA.Columns(1)
    cmd_Alterar.Enabled = True
    End If

End Sub




Private Sub txtCodigo_GotFocus()
txtCodigo.SelStart = 0
txtCodigo.SelLength = 100
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDescricao_GotFocus()
txtDescricao.SelStart = 0
txtDescricao.SelLength = 100
End Sub


Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub


Private Sub txtDescricao_LostFocus()
txtDescricao.Text = UCase(SemAcento(txtDescricao.Text))
End Sub

Private Sub txtSigla_GotFocus()
txtSigla.SelStart = 0
txtSigla.SelLength = 10
End Sub


Private Sub txtSigla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub


Private Sub txtSigla_LostFocus()
txtSigla.Text = UCase(txtSigla.Text)
End Sub

Private Sub txtUF_GotFocus()
txtUF.SelStart = 0
txtUF.SelLength = 10
End Sub


Private Sub txtUF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub


Private Sub txtUF_LostFocus()
txtUF.Text = UCase(txtUF.Text)
End Sub



