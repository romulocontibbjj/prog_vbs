VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaCLI 
   Caption         =   "Busca Clientes Remetente / Destinatários"
   ClientHeight    =   4515
   ClientLeft      =   1995
   ClientTop       =   1680
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   7965
   Begin VB.Frame fraConsCli 
      Caption         =   "Consulta Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4305
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7725
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   1920
         TabIndex        =   9
         Top             =   3480
         Width           =   3975
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2160
            TabIndex        =   6
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdTransportar 
            Caption         =   "Transportar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtBuscaNome 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   1
         Top             =   3045
         Width           =   2265
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   330
         Left            =   6480
         TabIndex        =   4
         Top             =   3000
         Width           =   1065
      End
      Begin VB.OptionButton optBuscaInic 
         Caption         =   "Busca no Início do Texto"
         Height          =   195
         Left            =   4080
         TabIndex        =   2
         Top             =   2940
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optBuscaTodo 
         Caption         =   "Busca no Texto Todo"
         Height          =   195
         Left            =   4080
         TabIndex        =   3
         Top             =   3150
         Width           =   2220
      End
      Begin MSDataGridLib.DataGrid GridConsCli 
         Bindings        =   "frmBuscaCLI.frx":0000
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   4471
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cgc"
            Caption         =   "cgc"
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
            DataField       =   "nome"
            Caption         =   "nome"
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
         BeginProperty Column02 
            DataField       =   "cidade"
            Caption         =   "cidade"
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
            BeginProperty Column00 
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3750,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1830,047
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Nome:"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   3045
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmBuscaCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBusca_Click()
    If txtBuscaNome.Text = "" Then
    Else
        If de_informa.rsSel_ConsCadCliNome.State = 1 Then de_informa.rsSel_ConsCadCliNome.Close
        If optBuscaInic = True Then
            de_informa.Sel_ConsCadCliNome Trim(txtBuscaNome) & "%"
            If de_informa.rsSel_ConsCadCliNome.RecordCount > 0 Then
                cmdTransportar.Enabled = True
            Else
                cmdTransportar.Enabled = False
            End If
        Else
            de_informa.Sel_ConsCadCliNome "%" & Trim(txtBuscaNome) & "%"
            If de_informa.rsSel_ConsCadCliNome.RecordCount > 0 Then
                cmdTransportar.Enabled = True
            Else
                cmdTransportar.Enabled = False
            End If
        End If
        GridConsCli.DataMember = "sel_consCadCliNome"
        GridConsCli.Refresh
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdTransportar_Click()
    Me.Hide
    If Me.Caption = "Busca Cliente REMETENTE" Then
        If frmPesquisaCTC.chkTodosEstab.Value = 0 Then
            frmPesquisaCTC.txtCGCRem = GridConsCli.Columns(0)
        Else
            frmPesquisaCTC.txtCGCRem = Mid(GridConsCli.Columns(0), 1, 8)
        End If
        frmPesquisaCTC.lblNomeRem = GridConsCli.Columns(1)
    ElseIf Me.Caption = "Busca Cliente DESTINATÁRIO" Then
        frmPesquisaCTC.TxtCGCDES = Mid(GridConsCli.Columns(0), 1, 8)
        frmPesquisaCTC.lblNomeDes = GridConsCli.Columns(1)
    ElseIf Me.Caption = "Busca Cliente REMETENTE - (Acompanhamento)" Then
        If frmAcompanha.chkTodosEstab.Value = 0 Then
            frmAcompanha.txtCGCRem = GridConsCli.Columns(0)
        Else
            frmAcompanha.txtCGCRem = Mid(GridConsCli.Columns(0), 1, 8)
        End If
        frmAcompanha.lblNomeRem = GridConsCli.Columns(1)
        frmAcompanha.optPer15d.SetFocus
    ElseIf Me.Caption = "Busca Cliente - Controle de Canhotos de NF" Then
        If frmGeraProtCanhotos.chkTodosEstab.Value = 0 Then
            frmGeraProtCanhotos.txtRemetCGC = GridConsCli.Columns(0)
        Else
            frmGeraProtCanhotos.txtRemetCGC = Mid(GridConsCli.Columns(0), 1, 8)
        End If
        frmGeraProtCanhotos.lblRemetNome = GridConsCli.Columns(1)
        frmGeraProtCanhotos.txtcopias.SetFocus
    ElseIf Me.Caption = "Busca Cliente Exporta EDI" Then
        If frmExportEDI.chkTodosEstab.Value = 0 Then
            frmExportEDI.txtCGCRem = GridConsCli.Columns(0)
        Else
            frmExportEDI.txtCGCRem = Mid(GridConsCli.Columns(0), 1, 8)
        End If
        frmExportEDI.lblNomeRem = GridConsCli.Columns(1)
        
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    txtBuscaNome.SetFocus
End Sub

Private Sub Form_Load()
    GridConsCli.DataMember = ""
    GridConsCli.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBuscaCLI = Nothing
End Sub

Private Sub GridConsCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        cmdTransportar_Click
    End If
End Sub

Private Sub optBuscaInic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optBuscaTodo_Click()
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtBuscaNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
