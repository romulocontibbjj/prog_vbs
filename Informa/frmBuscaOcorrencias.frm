VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaOcorrencias 
   Caption         =   "Pesquisa Cadastro de Ocorrências"
   ClientHeight    =   4260
   ClientLeft      =   2355
   ClientTop       =   1530
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7110
   Begin VB.Frame Frame9 
      Caption         =   "Cadastro de Ocorrências"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin VB.TextBox txtRefinar 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   2
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTransportar 
         Caption         =   "Transportar Selecionado"
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   3240
         Width           =   975
      End
      Begin VB.OptionButton optDescricao 
         Caption         =   "Por Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Por Código"
         Height          =   195
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid gridCadOcorr 
         Bindings        =   "frmBuscaOcorrencias.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8388608
         ForeColor       =   8454143
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
         DataMember      =   "Sel_BuscaOcorrDescr"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cod_ocorr"
            Caption         =   "Cód."
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
            Caption         =   "Descrição da Ocorrência"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
               Locked          =   -1  'True
               ColumnWidth     =   464,882
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   -1  'True
               Locked          =   -1  'True
               ColumnWidth     =   5504,882
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Refinar:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmBuscaOcorrencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub cmdTransportar_Click()
    frmPod.txtCodOcorr.Text = gridCadOcorr.Columns(0)
    Unload Me
End Sub
Private Sub Form_Load()
    If de_informa.rsSel_BuscaOcorrCod.State = 1 Then de_informa.rsSel_BuscaOcorrCod.Close
    de_informa.Sel_BuscaOcorrCod "%"
    
    If de_informa.rsSel_BuscaOcorrDescr.State = 1 Then de_informa.rsSel_BuscaOcorrDescr.Close
    de_informa.Sel_BuscaOcorrDescr "%"
    
    gridCadOcorr.DataMember = "Sel_BuscaOcorrDescr"
    gridCadOcorr.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuscaOcorrencias = Nothing
End Sub
Private Sub gridCadOcorr_Click()
    cmdTransportar.Enabled = True
End Sub

Private Sub gridCadOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdTransportar.SetFocus
    End If
End Sub

Private Sub optCodigo_Click()
    If de_informa.rsSel_BuscaOcorrCod.State = 1 Then de_informa.rsSel_BuscaOcorrCod.Close
    de_informa.Sel_BuscaOcorrCod "%"
    
    gridCadOcorr.DataMember = "Sel_BuscaOcorrCod"
    gridCadOcorr.Refresh
    cmdTransportar.Enabled = False
    txtRefinar.SetFocus
End Sub

Private Sub optDescricao_Click()
    If de_informa.rsSel_BuscaOcorrDescr.State = 1 Then de_informa.rsSel_BuscaOcorrDescr.Close
    de_informa.Sel_BuscaOcorrDescr "%"
    
    gridCadOcorr.DataMember = "Sel_BuscaOcorrDescr"
    gridCadOcorr.Refresh
    cmdTransportar.Enabled = False
    txtRefinar.SetFocus
End Sub
Private Sub txtRefinar_Change()
    If optDescricao = True Then
        If de_informa.rsSel_BuscaOcorrDescr.State = 1 Then de_informa.rsSel_BuscaOcorrDescr.Close
        de_informa.Sel_BuscaOcorrDescr "%" & Trim$(UCase(txtRefinar)) & "%"
        gridCadOcorr.DataMember = "Sel_BuscaOcorrDescr"
        'gridCadOcorr.Refresh
        If de_informa.rsSel_BuscaOcorrDescr.RecordCount = 1 Then
            cmdTransportar.Enabled = True
        Else
            cmdTransportar.Enabled = False
        End If
    ElseIf optCodigo = True Then
        If de_informa.rsSel_BuscaOcorrCod.State = 1 Then de_informa.rsSel_BuscaOcorrCod.Close
        de_informa.Sel_BuscaOcorrCod "%" & Trim$(UCase(txtRefinar)) & "%"
        gridCadOcorr.DataMember = "Sel_BuscaOcorrCod"
        'gridCadOcorr.Refresh
        If de_informa.rsSel_BuscaOcorrCod.RecordCount = 1 Then
            cmdTransportar.Enabled = True
        Else
            cmdTransportar.Enabled = False
        End If
    End If
End Sub

Private Sub txtRefinar_GotFocus()
    If optDescricao = True Then
        If de_informa.rsSel_BuscaOcorrDescr.RecordCount = 1 Then
            cmdTransportar.Enabled = True
        Else
            cmdTransportar.Enabled = False
        End If
    ElseIf optCodigo = True Then
        If de_informa.rsSel_BuscaOcorrCod.RecordCount = 1 Then
            cmdTransportar.Enabled = True
        Else
            cmdTransportar.Enabled = False
        End If
    End If
End Sub

Private Sub txtRefinar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
