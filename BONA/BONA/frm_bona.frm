VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_bona 
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   1545
   ClientTop       =   1260
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   12825
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      Begin MSDataGridLib.DataGrid grd_compara 
         Bindings        =   "frm_bona.frx":0000
         Height          =   1935
         Left            =   2160
         TabIndex        =   12
         Top             =   7320
         Width           =   7455
         _ExtentX        =   13150
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
         DataMember      =   "sel_compara"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "DATA"
            Caption         =   "DATA"
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
            DataField       =   "VALO_BONA"
            Caption         =   "VALO_BONA"
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
            DataField       =   "VALOR_INTEC"
            Caption         =   "VALOR_INTEC"
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
            DataField       =   "DIFERENCA"
            Caption         =   "DIFERENCA"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grd_tb_bona 
         Bindings        =   "frm_bona.frx":0017
         Height          =   1815
         Left            =   6240
         TabIndex        =   11
         Top             =   5400
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3201
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
         DataMember      =   "sel_bona"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "DATA"
            Caption         =   "DATA"
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
            DataField       =   "VALOR"
            Caption         =   "VALOR"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grd_fatura_valor 
         Bindings        =   "frm_bona.frx":002E
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   5400
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
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
         DataMember      =   "sel_fatura_valor"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "EMISSAO"
            Caption         =   "EMISSAO"
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
            DataField       =   "BRUTO_ICMS"
            Caption         =   "BRUTO_ICMS"
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
            DataField       =   "BRUTO"
            Caption         =   "BRUTO"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3615
         Begin VB.CommandButton cmd_buscar 
            Caption         =   "&Buscar"
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   840
            Width           =   1335
         End
         Begin MSMask.MaskEdBox mas_inicio 
            Height          =   300
            Left            =   240
            TabIndex        =   1
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mas_final 
            Height          =   300
            Left            =   1920
            TabIndex        =   2
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Caption         =   "Final:"
            Height          =   255
            Left            =   2160
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio:"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "á"
            Height          =   255
            Left            =   1560
            TabIndex        =   7
            Top             =   480
            Width           =   255
         End
      End
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid GRD_BONA 
         Bindings        =   "frm_bona.frx":0045
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   6165
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
         DataMember      =   "sel_fatura_periodo"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "FATURA"
            Caption         =   "FATURA"
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
            DataField       =   "EMISSAO"
            Caption         =   "EMISSAO"
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
            DataField       =   "VENC"
            Caption         =   "VENC"
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
            DataField       =   "CLIENTE"
            Caption         =   "CLIENTE"
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
            DataField       =   "VLBRUTOICMS"
            Caption         =   "VLBRUTOICMS"
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
            DataField       =   "VLBRUTO"
            Caption         =   "VLBRUTO"
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
            DataField       =   "ENVIADO"
            Caption         =   "ENVIADO"
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
            DataField       =   "ARQUIVO"
            Caption         =   "ARQUIVO"
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
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Line Line2 
         X1              =   6000
         X2              =   6000
         Y1              =   5400
         Y2              =   7200
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   12360
         Y1              =   5280
         Y2              =   5280
      End
   End
End
Attribute VB_Name = "frm_bona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_buscar_Click()

If deb_bona.rssel_fatura_periodo.State = 1 Then deb_bona.rssel_fatura_periodo.Close
    deb_bona.sel_fatura_periodo mas_inicio, mas_final
    
    If deb_bona.rssel_fatura_periodo.RecordCount < 1 Then
        MsgBox "Não há Faturas", vbInformation, "FAURAS"
        Exit Sub
    Else
        GRD_BONA.DataMember = "sel_fatura_periodo"
        GRD_BONA.Refresh
        
        
           
    End If
    
If deb_bona.rssel_fatura_valor.State = 1 Then deb_bona.rssel_fatura_valor.Close
   deb_bona.sel_fatura_valor mas_inicio, mas_final
    
    grd_fatura_valor.DataMember = "sel_fatura_valor"
    grd_fatura_valor.Refresh
    
If deb_bona.rssel_bona.State = 1 Then deb_bona.rssel_bona.Close
    deb_bona.sel_bona mas_inicio, mas_final
    
    grd_tb_bona.DataMember = "sel_bona"
    grd_tb_bona.Refresh
    
    
    
If deb_bona.rssel_compara.State = 1 Then deb_bona.rssel_compara.Close
    deb_bona.sel_compara mas_inicio, mas_final
    
    grd_compara.DataMember = "sel_compara"
    grd_compara.Refresh
    


    
    


End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub
