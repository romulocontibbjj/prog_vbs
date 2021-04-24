VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_busca_cliente 
   Caption         =   "Busca de Clientes INTEC Transportes"
   ClientHeight    =   5655
   ClientLeft      =   1605
   ClientTop       =   1965
   ClientWidth     =   6630
   ControlBox      =   0   'False
   Icon            =   "frm_busca_cliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   5100
      Width           =   1755
   End
   Begin VB.CommandButton cmd_seleciona 
      Caption         =   "Seleciona Cliente"
      Height          =   435
      Left            =   4740
      TabIndex        =   8
      Top             =   5100
      Width           =   1755
   End
   Begin VB.TextBox txt_proc_nome 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.TextBox txt_nome 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   4020
      Width           =   6375
   End
   Begin VB.TextBox txt_cgc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   4620
      Width           =   4215
   End
   Begin VB.TextBox txt_proc_cgc 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   4815
   End
   Begin VB.CommandButton cmd_busca_nome 
      Caption         =   "Busque!"
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmd_busca_cgc 
      Caption         =   "Busque!"
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   1020
      Width           =   1455
   End
   Begin VB.CheckBox chk_todos_est 
      Caption         =   "Todos Estabelecimentos"
      Height          =   195
      Left            =   4440
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid grid_clientes 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   8388608
      ColumnHeaders   =   -1  'True
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_proc_nome 
      Caption         =   "Digite o nome que você procura"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbl_proc_cgc 
      Caption         =   "Digite o CGC que você procura"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "CGC do Cliente"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   4380
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Nome do Cliente"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   3780
      Width           =   2415
   End
End
Attribute VB_Name = "frm_busca_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xMuda As Boolean

Private Sub cmd_Busca_CGC_Click()
cmd_busca_nome.Enabled = False
cmd_busca_cgc.Enabled = False

xMuda = False
    If de_informa.rsSel_BuscaCGC.State = 1 Then de_informa.rsSel_BuscaCGC.Close
    de_informa.Sel_BuscaCGC "%" & txt_proc_cgc.Text & "%"
    
    Set grid_clientes.DataSource = de_informa
    grid_clientes.DataMember = "sel_BuscaCGC"
    grid_clientes.Refresh
    If de_informa.rsSel_BuscaCGC.RecordCount > 0 Then xMuda = True
cmd_busca_nome.Enabled = True
cmd_busca_cgc.Enabled = True
End Sub

Private Sub cmd_Busca_Nome_Click()
cmd_busca_nome.Enabled = False
cmd_busca_cgc.Enabled = False
xMuda = False

    If de_informa.rsSel_BuscaNome.State = 1 Then de_informa.rsSel_BuscaNome.Close
    de_informa.Sel_BuscaNome "%" & txt_proc_nome.Text & "%"
    
    Set grid_clientes.DataSource = de_informa
    grid_clientes.DataMember = "sel_BuscaNome"
    grid_clientes.Refresh
    
    If de_informa.rsSel_BuscaNome.RecordCount > 0 Then xMuda = True

cmd_busca_cgc.Enabled = True
cmd_busca_nome.Enabled = True
End Sub

Private Sub cmd_cancelar_Click()
Unload Me
End Sub

Private Sub cmd_seleciona_Click()
cmd_seleciona.Enabled = False

    If Len(txt_nome.Text) = 0 Or Len(txt_cgc.Text) = 0 Then
    MsgBox "Você não tem todas as informações necessárias para poder sair...", vbExclamation, ""
    cmd_seleciona.Enabled = True
    Exit Sub
    End If

xForm.TxtNomeCliente.Text = txt_nome.Text

    If chk_todos_est.Value = 1 Then
    xForm.TxtCGCCliente.Text = Mid(txt_cgc.Text, 1, 8)
    Else
    xForm.TxtCGCCliente.Text = txt_cgc.Text
    End If

Unload Me
End Sub

Private Sub Form_Activate()
txt_proc_nome.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set xForm = Nothing
Set ActiveForm = Nothing
End Sub

Private Sub grid_clientes_Click()
    If xMuda = True Then
    txt_nome.Text = grid_clientes.Columns(1)
    txt_cgc.Text = grid_clientes.Columns(0)
    ElseIf xMuda = False Then
    txt_nome.Text = ""
    txt_cgc.Text = ""
    End If
End Sub

Private Sub grid_clientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
grid_clientes_Click
End Sub

Private Sub txt_proc_cgc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If
End Sub

Private Sub txt_proc_nome_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If
End Sub
