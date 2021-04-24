VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMotivosDesconto 
   Caption         =   "Motivos de Desconto de Fatura"
   ClientHeight    =   3180
   ClientLeft      =   2430
   ClientTop       =   1200
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4725
   Begin MSDataGridLib.DataGrid gridMotivos 
      Bindings        =   "frmMotivosDesconto.frx":0000
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
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
      DataMember      =   "Sel_MotivosDesc"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "motivo"
         Caption         =   "motivo"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3869,858
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<ENTER>=Escolher"
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
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "frmMotivosDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If de_informa.rsSel_MotivosDesc.State = 1 Then de_informa.rsSel_MotivosDesc.Close
    de_informa.Sel_MotivosDesc
    
    gridMotivos.DataMember = "sel_motivosdesc"
    gridMotivos.Refresh
    
End Sub

Private Sub gridMotivos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Me.Caption = "Conceder Desconto" Then
            frmDescontos.lblTipoAbat = gridMotivos.Columns(1)
        Else
            frmGravaNovaFatura.lblTipoAbat = gridMotivos.Columns(1)
        End If
        Unload Me
    End If
    If KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
End Sub
