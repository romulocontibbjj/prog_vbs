VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProcBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procura de Bancos"
   ClientHeight    =   4290
   ClientLeft      =   2985
   ClientTop       =   2940
   ClientWidth     =   6555
   ControlBox      =   0   'False
   Icon            =   "frmProcBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6555
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "Continuar  >>"
      Height          =   375
      Left            =   5100
      TabIndex        =   3
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "<<  Voltar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3780
      Width           =   1335
   End
   Begin VB.TextBox txtNomeBanco 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   3300
      Width           =   5115
   End
   Begin VB.TextBox txtNumBanco 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   3300
      Width           =   1035
   End
   Begin VB.CommandButton cmd_busca 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5100
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtProcBanco 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   4875
   End
   Begin MSDataGridLib.DataGrid GridBco 
      Bindings        =   "frmProcBancos.frx":000C
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3413
      _Version        =   393216
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
      DataMember      =   "Sel_CadBco"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "bconum"
         Caption         =   "Nº Banco"
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
         DataField       =   "bconome"
         Caption         =   "Nome do Banco"
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
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004,788
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nº Banco"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3060
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome Banco"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   3060
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Procura pelo Nome do Banco"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   2100
   End
End
Attribute VB_Name = "frmProcBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_busca_Click()
If de_informa.rsSel_CadBco.State = 1 Then de_informa.rsSel_CadBco.Close
de_informa.Sel_CadBco "%" & Trim(txtProcBanco.Text) & "%"

GridBco.DataMember = "Sel_CadBco"
GridBco.Refresh
DoEvents

    If de_informa.rsSel_CadBco.RecordCount = 0 Then
    GridBco.Enabled = False
    Else
    GridBco.Enabled = True
    End If
    
End Sub

Private Sub cmdContinuar_Click()
    If Len(Trim(txtNumBanco.Text)) = 0 Or Len(Trim(txtNomeBanco.Text)) = 0 Then
    MsgBox "Você não selecionou Banco algum... Por favor, tente novamente", vbExclamation, ""
    Exit Sub
    Else
    xForm.txtNomeBanco.Text = txtNomeBanco.Text
    xForm.txtNumBanco.Text = txtNumBanco.Text
    Unload Me
    End If
End Sub

Private Sub cmdVoltar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtProcBanco.SetFocus
End Sub

Private Sub Form_Load()
    If de_informa.rsSel_CadBco.State = 1 Then de_informa.rsSel_CadBco.Close
    de_informa.Sel_CadBco "%"
    GridBco.DataMember = "Sel_CadBco"
    GridBco.Refresh
    
    If de_informa.rsSel_CadBco.RecordCount = 0 Then
    GridBco.Enabled = False
    Else
    GridBco.Enabled = True
    End If
    
    txtNomeBanco.Text = ""
    txtNumBanco.Text = ""
End Sub


Private Sub GridBco_Click()
txtNomeBanco.Text = GridBco.Columns(1)
txtNumBanco.Text = GridBco.Columns(0)
DoEvents
End Sub

Private Sub GridBco_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
txtNomeBanco.Text = GridBco.Columns(1)
txtNumBanco.Text = GridBco.Columns(0)
DoEvents
End Sub


Private Sub txtProcBanco_GotFocus()
txtProcBanco.SelStart = 0
txtProcBanco.SelLength = 300
End Sub

Private Sub txtProcBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
