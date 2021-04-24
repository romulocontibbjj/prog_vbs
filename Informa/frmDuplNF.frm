VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDuplNF 
   Caption         =   "Número de NFs Duplicadas"
   ClientHeight    =   3555
   ClientLeft      =   1320
   ClientTop       =   1890
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   9255
   Begin VB.Frame Frame1 
      Caption         =   "Nota Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   1905
      Begin VB.Label lblNfDupl 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   4
         Top             =   315
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Seleciona"
      Height          =   435
      Left            =   6195
      TabIndex        =   2
      Top             =   315
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   7770
      TabIndex        =   1
      Top             =   315
      Width           =   1380
   End
   Begin MSDataGridLib.DataGrid gridDuplNF 
      Bindings        =   "frmDuplNF.frx":0000
      Height          =   2430
      Left            =   105
      TabIndex        =   0
      Top             =   945
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4286
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
      DataMember      =   "Sel_DuplNF"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "serie"
         Caption         =   "Série"
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
         DataField       =   "filialctc"
         Caption         =   "Filial-CTC"
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
         DataField       =   "data"
         Caption         =   "Data"
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
      BeginProperty Column03 
         DataField       =   "tem_ocorr"
         Caption         =   "St."
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
      BeginProperty Column04 
         DataField       =   "remet_nome"
         Caption         =   "Remetente"
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
      BeginProperty Column05 
         DataField       =   "dest_nome"
         Caption         =   "Destinatário"
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
      BeginProperty Column06 
         DataField       =   "cidade_dest"
         Caption         =   "Cidade Destino"
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
      BeginProperty Column07 
         DataField       =   "uf_dest"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   269,858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2369,764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2654,929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   345,26
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDuplNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Me.Hide
    If Me.Caption = "SAC - Número de NFs Duplicadas" Then
        frmSac.txtNumNf.SetFocus
    ElseIf Me.Caption = "POD - Número de NFs Duplicadas" Then
        frmPod.txtNumNf.SetFocus
    End If
    Unload frmDuplNF
End Sub
Private Sub CmdOk_Click()
    Me.Hide
    If Me.Caption = "SAC - Número de NFs Duplicadas" Then
        frmSac.optCTC = True
        frmSac.TxtFilial.Text = Mid(gridDuplNF.Columns(1), 1, 2)
        frmSac.txtCtc.Text = Mid(gridDuplNF.Columns(1), 3, 8)
        DoEvents
        frmSac.cmbProcurar.SetFocus
    ElseIf Me.Caption = "POD - Número de NFs Duplicadas" Then
        frmPod.optCTC = True
        frmPod.TxtFilial.Text = Mid(gridDuplNF.Columns(1), 1, 2)
        frmPod.txtCtc.Text = Mid(gridDuplNF.Columns(1), 3, 8)
        frmPod.cmdProcurar.SetFocus
    End If
    DoEvents
    SendKeys "{ENTER}"
    DoEvents
    Unload frmDuplNF
End Sub
Private Sub Form_Activate()
    If de_informa.rsSel_DuplNF.State = 1 Then de_informa.rsSel_DuplNF.Close
    If Me.Caption = "SAC - Número de NFs Duplicadas" Then
        de_informa.Sel_DuplNF Val(frmSac.txtNumNf)
    ElseIf Me.Caption = "POD - Número de NFs Duplicadas" Then
        de_informa.Sel_DuplNF Val(frmPod.txtNumNf)
    End If
    If de_informa.rsSel_DuplNF.RecordCount = 0 Then
        MsgBox "Erro de Consistência. Chame Suporte Técnico !"
        Exit Sub
    Else
        'preenche o grid
        gridDuplNF.DataMember = "Sel_DuplNF"
        gridDuplNF.Refresh
    End If
    If Me.Caption = "SAC - Número de NFs Duplicadas" Then
        lblNfDupl.Caption = Val(frmSac.txtNumNf)
    ElseIf Me.Caption = "POD - Número de NFs Duplicadas" Then
        lblNfDupl.Caption = Val(frmPod.txtNumNf)
    End If
End Sub
Private Sub Form_Load()
    'MsgBox "Este Número de NF existe mais de uma vez no Banco de Dados. A seguir, escolha a NF que deseja consultar.", vbCritical + vbOKOnly, "NF Duplicada"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDuplNF = Nothing
End Sub

Private Sub gridDuplNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        CmdOk_Click
    ElseIf KeyAscii = 27 Then   'TECLA ESC
        KeyAscii = 0
        Unload Me
    End If
End Sub
