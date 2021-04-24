VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCadLocalidadeCONF 
   Caption         =   "Clique sobre a linha da Cidade desejada e pressione Ok."
   ClientHeight    =   4995
   ClientLeft      =   2250
   ClientTop       =   1365
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "FrmCadLocalidadeCONF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   6480
      TabIndex        =   5
      Top             =   2160
      Width           =   315
   End
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "CONTINUAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox TxtUF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Top             =   4560
      Width           =   675
   End
   Begin VB.TextBox TxtCidade 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   5835
   End
   Begin MSDataGridLib.DataGrid GridCidades 
      Bindings        =   "FrmCadLocalidadeCONF.frx":000C
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7117
      _Version        =   393216
      ForeColor       =   -2147483630
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
      DataMember      =   "Sel_CONFCidade"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cidade"
         Caption         =   "Cidade"
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
         DataField       =   "uf"
         Caption         =   "UF"
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
            ColumnWidth     =   4995,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   540,284
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "UF"
      Height          =   195
      Left            =   6120
      TabIndex        =   6
      Top             =   4320
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4260
      Width           =   495
   End
End
Attribute VB_Name = "FrmCadLocalidadeCONF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
LeaveSub = True
Unload Me
End Sub

Private Sub cmdContinuar_Click()
cmdContinuar.Enabled = False
CmdCancelar.Enabled = False
    
    If Len(Trim(txtCidade.Text)) = 0 Then Exit Sub

xForm.txtLocalidade.Text = txtCidade.Text
xForm.txtUF.Text = txtUF.Text
Unload Me

End Sub

Private Sub Form_Load()
If de_informa.rsSel_CONFCidade.State = 1 Then de_informa.rsSel_CONFCidade.Close
de_informa.Sel_CONFCidade "%" & Trim(xForm.txtLocalidade.Text) & "%"

GridCidades.DataMember = "Sel_CONFCidade"
GridCidades.Refresh
End Sub

Private Sub GridCidades_Click()
txtCidade.Text = GridCidades.Columns(0)
txtUF.Text = GridCidades.Columns(1)
End Sub

Private Sub GridCidades_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GridCidades_Click
End Sub
