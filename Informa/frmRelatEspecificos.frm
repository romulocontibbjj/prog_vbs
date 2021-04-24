VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRelatEspecificos 
   Caption         =   "Gera Arquivos/RelatÛrios EspecÌficos"
   ClientHeight    =   8010
   ClientLeft      =   690
   ClientTop       =   915
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   11775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexResultado 
         Height          =   2655
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4683
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.OptionButton optTxt 
         Caption         =   "TXT"
         Height          =   195
         Left            =   9840
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         Height          =   195
         Left            =   10680
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdExporta 
         Caption         =   "Exportar Arquivo ..."
         Height          =   350
         Left            =   7200
         TabIndex        =   11
         Top             =   200
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc ado_informa 
         Height          =   330
         Left            =   1200
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   60
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=intec;Data Source=."
         OLEDBString     =   "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=intec;Data Source=."
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select top 1 * from tb_cadcli"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblTotLinha 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   6120
         TabIndex        =   46
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   6000
         TabIndex        =   45
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblLinha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   5880
         TabIndex        =   44
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Par‚metros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   3150
      Width           =   11775
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   10560
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   10560
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPara8 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara5 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara7 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara6 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara4 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara2 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8520
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara1 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2330
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPer1 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblPara8 
         AutoSize        =   -1  'True
         Caption         =   "Para8...................:"
         Height          =   195
         Left            =   7080
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara5 
         AutoSize        =   -1  'True
         Caption         =   "Para5...................:"
         Height          =   195
         Left            =   7080
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara7 
         AutoSize        =   -1  'True
         Caption         =   "Para7...................:"
         Height          =   195
         Left            =   3600
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara6 
         AutoSize        =   -1  'True
         Caption         =   "Para6...................:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara4 
         AutoSize        =   -1  'True
         Caption         =   "Para4...................:"
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara3 
         AutoSize        =   -1  'True
         Caption         =   "Para3...................:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara2 
         AutoSize        =   -1  'True
         Caption         =   "Para2...................:"
         Height          =   195
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPara1 
         AutoSize        =   -1  'True
         Caption         =   "Para1...................:"
         Height          =   195
         Left            =   3600
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblPeriodo 
         AutoSize        =   -1  'True
         Caption         =   "PerÌodo:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblPeriodoa 
         AutoSize        =   -1  'True
         Caption         =   "‡"
         Height          =   195
         Left            =   2120
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RelatÛrios DisponÌveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   11775
      Begin MSDataGridLib.DataGrid gridRelatorios 
         Bindings        =   "frmRelatEspecificos.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4048
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
         DataMember      =   "Sel_Relatorios"
         ColumnCount     =   18
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
            DataField       =   "descricao"
            Caption         =   "descricao"
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
            DataField       =   "usuario"
            Caption         =   "usuario"
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
            DataField       =   "tipo"
            Caption         =   "tipo"
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
            DataField       =   "senha"
            Caption         =   "senha"
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
            DataField       =   "querytext"
            Caption         =   "querytext"
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
            DataField       =   "parametros"
            Caption         =   "parametros"
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
         BeginProperty Column08 
            DataField       =   "paraperiodo"
            Caption         =   "paraperiodo"
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
         BeginProperty Column09 
            DataField       =   "para1"
            Caption         =   "para1"
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
         BeginProperty Column10 
            DataField       =   "para2"
            Caption         =   "para2"
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
         BeginProperty Column11 
            DataField       =   "para3"
            Caption         =   "para3"
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
         BeginProperty Column12 
            DataField       =   "para4"
            Caption         =   "para4"
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
         BeginProperty Column13 
            DataField       =   "para5"
            Caption         =   "para5"
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
         BeginProperty Column14 
            DataField       =   "para6"
            Caption         =   "para6"
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
         BeginProperty Column15 
            DataField       =   "para7"
            Caption         =   "para7"
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
         BeginProperty Column16 
            DataField       =   "para8"
            Caption         =   "para8"
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
         BeginProperty Column17 
            DataField       =   "datacad"
            Caption         =   "datacad"
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
               ColumnWidth     =   555,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5009,953
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Caption         =   "DescriÁ„o do RelatÛrio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   6360
         TabIndex        =   28
         Top             =   510
         Width           =   5295
         Begin VB.CheckBox chkMostraQuery 
            Caption         =   "Mostrar Query"
            Height          =   225
            Left            =   3840
            TabIndex        =   35
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluirRelat 
            Caption         =   "Novo ..."
            Enabled         =   0   'False
            Height          =   420
            Left            =   4080
            TabIndex        =   39
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditRelat 
            Caption         =   "Detalhar ..."
            Enabled         =   0   'False
            Height          =   420
            Left            =   2640
            TabIndex        =   40
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtNumRelat 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   41
            Top             =   1860
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "ID do RelatÛrio:"
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
            TabIndex        =   42
            Top             =   1905
            Width           =   1365
         End
         Begin VB.Label lblDescricao 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label lblCadastro 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3480
            TabIndex        =   33
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblUsuario 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   840
            TabIndex        =   32
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Data Cadastro:"
            Height          =   195
            Left            =   2280
            TabIndex        =   31
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Usu·rio:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label lblDescrQuery 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   5055
         End
      End
      Begin VB.OptionButton optTodosRelat 
         Caption         =   "Todos os RelatÛrios"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   300
         Width           =   1815
      End
      Begin VB.OptionButton optUsuarioRelat 
         Caption         =   "Somente do Usu·rio:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmRelatEspecificos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelWS1 As Excel.Worksheet

Private Sub chkMostraQuery_Click()
    If chkMostraQuery.Value = 1 Then
        lblDescrQuery.Visible = True
        lblDescricao.Visible = False
        lblUsuario.Visible = False
        lblCadastro.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        Label15.Visible = False
        txtNumRelat.Visible = False
        cmdEditRelat.Visible = False
        cmdIncluirRelat.Visible = False
    Else
        lblDescrQuery.Visible = False
        lblDescricao.Visible = True
        lblUsuario.Visible = True
        lblCadastro.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        Label15.Visible = True
        txtNumRelat.Visible = True
        cmdEditRelat.Visible = True
        cmdIncluirRelat.Visible = True
    End If
End Sub
Private Sub cmdExporta_Click()
    Dim xfile As String, xfile0 As String

    cmdExporta.Caption = "Aguarde..."
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Me.MousePointer = 11
    DoEvents
    
    xfile0 = InputBox("Entre com o Nome do Arquivo (sem extens„o):")
    xfile = ""
    
    If Len(xfile0) > 40 Then xfile0 = Mid$(xfile0, 1, 40)

    For X = 1 To Len(xfile0)
        If InStr(1, " QWERTYUIOPASDFGHJKLZXCVBNM-_1234567890¡…Õ”⁄√’¿¬ ‘«", UCase(Mid$(xfile0, X, 1)), vbTextCompare) = 0 Then
        Else
            xfile = xfile + Mid$(xfile0, X, 1)
        End If
    Next
    
    xfile = Trim$(xfile)
    
    If Len(xfile) < 1 Then
        MsgBox "Nome de Arquivo Inv·lido !!!", vbCritical, "ERRO"
        cmdExporta.Caption = "Exportar Arquivo ..."
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        Me.MousePointer = 0
        DoEvents
        Exit Sub
    End If
    
    If optTxt.Value = True Then
        xfile = "c:\informa\" & xfile & ".txt"
    ElseIf optExcel = True Then
        xfile = "c:\informa\" & xfile & ".xls"
    End If
    
    If Len(Trim$(Dir(xfile))) > 0 Then
        MsgBox "J· Existe Um Arquivo Com Este Nome. Escolha Novamente <Exportar Arquivo ...> Digite Outro Nome !", vbCritical, "ERRO"
        cmdExporta.Caption = "Exportar Arquivo ..."
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        Me.MousePointer = 0
        DoEvents
        Exit Sub
    End If
    
    If optTxt.Value = True Then
        
        Open xfile For Output As #1
        
        For X = 0 To FlexResultado.Rows - 1
            xlinha = ""
            lblLinha = X + 1
            lblTotLinha = FlexResultado.Rows - 1
            DoEvents
            For Y = 1 To FlexResultado.Cols - 1
                xlinha = xlinha & FlexResultado.TextMatrix(X, Y) & "#"
            Next
            Print #1, xlinha
        Next
        
        lblLinha = lblLinha - 1
        
        Close #1
        
        Me.MousePointer = 0
        MsgBox "Arquivo TXT Gerado ! Para Import·-lo para o Excel, Abra-o como Arquivo Texto e Escolha como Delimitador o caracter '#' ." + Chr(13) + Chr(10) + Chr(13) + Chr(10) + xfile, vbInformation, "Arquivo Gerado"
        
    ElseIf optExcel.Value = True Then
    
        Set Excel = CreateObject("Excel.Application")
        Set Excel = GetObject(, "Excel.Application")
        Excel.Visible = False
        Set ExcelWBk = Excel.Workbooks.Add
        Set ExcelWS1 = ExcelWBk.Worksheets(1)
    
        ExcelWS1.Cells.Font.Name = "Verdana"
        
        For X = 0 To FlexResultado.Rows - 1
            xlinha = ""
            lblLinha = X + 1
            lblTotLinha = FlexResultado.Rows - 1
            DoEvents
            For Y = 1 To FlexResultado.Cols - 1
                ExcelWS1.Cells(X + 5, Y) = FlexResultado.TextMatrix(X, Y)
            Next
        Next
        
        lblLinha = lblLinha - 1
        
        ExcelWS1.Range(ExcelWS1.Cells(5, 1), ExcelWS1.Cells(5, FlexResultado.Cols - 1)).Font.Bold = True
        ExcelWS1.Range(ExcelWS1.Cells(5, 1), ExcelWS1.Cells(5, FlexResultado.Cols - 1)).HorizontalAlignment = xlCenter
        ExcelWS1.Range(ExcelWS1.Cells(5, 1), ExcelWS1.Cells(5, FlexResultado.Cols - 1)).VerticalAlignment = xlCenter
        ExcelWS1.Range(ExcelWS1.Cells(5, 1), ExcelWS1.Cells(FlexResultado.Rows + 4, FlexResultado.Cols - 1)).Borders.ColorIndex = 1
        ExcelWS1.Range("A:DZ").EntireColumn.AutoFit
        ExcelWS1.Cells(1, 1) = gridRelatorios.Columns(2)
        
        xpara = ""
        
        If lblPeriodo.Visible = True Then
            ExcelWS1.Cells(2, 1) = "PERÕODO: " & mskPer1 & " a " & mskPer2
        End If
        If lblPara1.Visible = True Then
            xpara = xpara + lblPara1 & " " & txtPara1 & " / "
        End If
        If lblPara2.Visible = True Then
            xpara = xpara + lblPara2 & " " & txtPara2 & " / "
        End If
        If lblPara3.Visible = True Then
            xpara = xpara + lblPara3 & " " & txtPara3 & " / "
        End If
        If lblPara4.Visible = True Then
            xpara = xpara + lblPara4 & " " & txtPara4 & " / "
        End If
        If lblPara5.Visible = True Then
            xpara = xpara + lblPara5 & " " & txtPara5 & " / "
        End If
        If lblPara6.Visible = True Then
            xpara = xpara + lblPara6 & " " & txtPara6 & " / "
        End If
        If lblPara7.Visible = True Then
            xpara = xpara + lblPara7 & " " & txtPara7 & " / "
        End If
        If lblPara8.Visible = True Then
            xpara = xpara + lblPara8 & " " & txtPara8 & " / "
        End If
        
        If Len(xpara) > 1 Then xpara = "PAR¬METROS: " & Trim$(Mid$(xpara, 1, Len(xpara) - 2))
        
        ExcelWS1.Cells(3, 1) = xpara
        ExcelWS1.Range(ExcelWS1.Cells(1, 1), ExcelWS1.Cells(3, 1)).Font.Bold = True
        ExcelWBk.SaveAs xfile, , , , , , xlExclusive
        
        ExcelWBk.Close
        
        Me.MousePointer = 0
        MsgBox "Arquivo Formato Excel Gerado !" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + xfile, vbInformation, "Arquivo Gerado"
        
    End If
        
    cmdExporta.Caption = "Exportar Arquivo ..."
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    DoEvents
    
    
End Sub
Private Sub cmdGerar_Click()
    Dim xquery As String, xcaract As Integer, xcontrol As Integer
    Dim rsVar As Variant
    
    lblLinha = "0"
    lblTotLinha = "0"
    
    cmdGerar.Caption = "Aguarde..."
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Me.MousePointer = 11
    DoEvents
    
    FlexResultado.Clear
    ado_informa.RecordSource = "select filialctc from tb_ctc_esp where filialctc = 'xxxxxxxxxx'"
    ado_informa.Refresh
    Set FlexResultado.DataSource = ado_informa
    FlexResultado.Refresh
    
    xquery = lblDescrQuery.Caption
    
    If mskPer1.Visible = True Then
        xcontrol = -2
    Else
        xcontrol = 0
    End If
    
    Do While InStr(1, xquery, "?", vbTextCompare) > 0
        xcontrol = xcontrol + 1
        If xcontrol = -1 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & (Mid$(mskPer1, 7, 4) & "/" & Mid$(mskPer1, 4, 2) & "/" & Mid$(mskPer1, 1, 2)) & "'" & _
                     Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 0 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & (Mid$(mskPer2, 7, 4) & "/" & Mid$(mskPer2, 4, 2) & "/" & Mid$(mskPer2, 1, 2)) & "'" & _
                     Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 1 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara1 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 2 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara2 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 3 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara3 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 4 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara4 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 5 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara5 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 6 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara6 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 7 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara7 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        ElseIf xcontrol = 8 Then
            xquery = Mid$(xquery, 1, InStr(1, xquery, "?", vbTextCompare) - 1) & "'" & txtPara8 & "'" & Mid$(xquery, InStr(1, xquery, "?", vbTextCompare) + 1)
        End If
    Loop

    ado_informa.RecordSource = xquery
    ado_informa.Refresh
    
    FlexResultado.Clear
    
    FlexResultado.Rows = ado_informa.Recordset.RecordCount + 1
    FlexResultado.Cols = ado_informa.Recordset.Fields.Count + 1
    lblTotLinha = Trim$(Str(ado_informa.Recordset.RecordCount))

    'Faz as Colunas da Flex
    For I = 0 To ado_informa.Recordset.Fields.Count - 1
        FlexResultado.TextMatrix(0, I + 1) = ado_informa.Recordset.Fields(I).Name
    Next
    
    If ado_informa.Recordset.RecordCount > 0 Then
    
        ado_informa.Recordset.MoveFirst
        xLin = 0
        Do Until ado_informa.Recordset.EOF
            xLin = xLin + 1
            'Preenche a Grid da Flex com os Dados
            For I = 0 To ado_informa.Recordset.Fields.Count - 1
                If Not IsNull(ado_informa.Recordset.Fields(I)) Then
                    If IsDate(ado_informa.Recordset.Fields(I)) Then
                        FlexResultado.TextMatrix(xLin, I + 1) = Trim$(Str(Year(ado_informa.Recordset.Fields(I)))) + "/" + _
                                                                zeros2(Str(Month(ado_informa.Recordset.Fields(I))), 2) + "/" + _
                                                                zeros2(Str(Day(ado_informa.Recordset.Fields(I))), 2)
                    ElseIf VarType(ado_informa.Recordset.Fields(I)) = vbCurrency Then
                        xvalor = Trim$(SoNumeros(Format(ado_informa.Recordset.Fields(I), "##,###,##0.00")))
                        FlexResultado.TextMatrix(xLin, I + 1) = Mid$(xvalor, 1, Len(xvalor) - 2) & "." & Mid$(xvalor, Len(xvalor) - 1)
                    Else
                        FlexResultado.TextMatrix(xLin, I + 1) = ado_informa.Recordset.Fields(I)
                    End If
                End If
            Next
            ado_informa.Recordset.MoveNext
        Loop
        
    End If

'    FlexResultado.Row = 1
'    FlexResultado.Col = 1

    ' Set range of cells in the grid
'    FlexResultado.RowSel = FlexResultado.Rows - 1
'    FlexResultado.ColSel = FlexResultado.Cols - 1
'    FlexResultado.Clip = rsVar

    ' Reset the grid's selected range of cells
'    FlexResultado.RowSel = FlexResultado.Row
'    FlexResultado.ColSel = FlexResultado.Col
    
'    Set FlexResultado.DataSource = ado_informa
   ' FlexResultado.Refresh
    
    Me.MousePointer = 0
    DoEvents
    
    If ado_informa.Recordset.RecordCount < 1 Then
        MsgBox "N„o Foram Encontrados Dados Com Estes Par‚metros !!", vbCritical
    End If
    
    cmdGerar.Caption = "Gerar"
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    DoEvents
    
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    FlexResultado.ColWidth(0) = 300
    
    If de_informa.rsSel_Relatorios.State = 1 Then de_informa.rsSel_Relatorios.Close
    de_informa.Sel_Relatorios xusuario
    
    optUsuarioRelat.Caption = optUsuarioRelat.Caption & " " & xusuario
    
    gridRelatorios.DataMember = "Sel_relatorios"
    
    ado_informa.ConnectionString = xstrcon
    
End Sub

Private Sub gridRelatorios_Click()
    lblDescricao = gridRelatorios.Columns(2)
    lblDescrQuery = gridRelatorios.Columns(6)
    lblUsuario = gridRelatorios.Columns(3)
    lblCadastro = gridRelatorios.Columns(17)
    txtNumRelat = gridRelatorios.Columns(0)
    
    If Len(Trim$(gridRelatorios.Columns(8))) = 0 Then
        lblPeriodo.Visible = False
        lblPeriodoa.Visible = False
        mskPer1.Visible = False
        mskPer2.Visible = False
    Else
        lblPeriodo.Visible = True
        lblPeriodoa.Visible = True
        mskPer1.Visible = True
        mskPer2.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(9))) = 0 Then
        lblPara1.Visible = False
        txtPara1.Visible = False
    Else
        lblPara1.Visible = True
        lblPara1.Caption = gridRelatorios.Columns(9) & ":"
        txtPara1.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(10))) = 0 Then
        lblPara2.Visible = False
        txtPara2.Visible = False
    Else
        lblPara2.Visible = True
        lblPara2.Caption = gridRelatorios.Columns(10) & ":"
        txtPara2.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(11))) = 0 Then
        lblPara3.Visible = False
        txtPara3.Visible = False
    Else
        lblPara3.Visible = True
        lblPara3.Caption = gridRelatorios.Columns(11) & ":"
        txtPara3.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(12))) = 0 Then
        lblPara4.Visible = False
        txtPara4.Visible = False
    Else
        lblPara4.Visible = True
        lblPara4.Caption = gridRelatorios.Columns(12) & ":"
        txtPara4.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(13))) = 0 Then
        lblPara5.Visible = False
        txtPara5.Visible = False
    Else
        lblPara5.Visible = True
        lblPara5.Caption = gridRelatorios.Columns(13) & ":"
        txtPara5.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(14))) = 0 Then
        lblPara6.Visible = False
        txtPara6.Visible = False
    Else
        lblPara6.Visible = True
        lblPara6.Caption = gridRelatorios.Columns(14) & ":"
        txtPara6.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(15))) = 0 Then
        lblPara7.Visible = False
        txtPara7.Visible = False
    Else
        lblPara7.Visible = True
        lblPara7.Caption = gridRelatorios.Columns(15) & ":"
        txtPara7.Visible = True
    End If
    If Len(Trim$(gridRelatorios.Columns(16))) = 0 Then
        lblPara8.Visible = False
        txtPara8.Visible = False
    Else
        lblPara8.Visible = True
        lblPara8.Caption = gridRelatorios.Columns(16) & ":"
        txtPara8.Visible = True
    End If
    
    cmdGerar.Enabled = True
    
End Sub

Private Sub mskPer1_GotFocus()
    mskPer1.SelStart = 0
    mskPer1.SelLength = 10
End Sub

Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer1_LostFocus()
    If mskPer1.Text <> "__/__/____" Then
        mskPer1.Text = century(mskPer1.Text)
        If IsDate(mskPer1.Text) = False Or Mid(mskPer1.Text, 4, 2) > 12 Then
            MsgBox "Data Inv·lida !", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
        If CDate(mskPer1.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub mskPer2_GotFocus()
    mskPer2.SelStart = 0
    mskPer2.SelLength = 10
End Sub

Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer2_LostFocus()
    If mskPer2.Text <> "__/__/____" Then
        mskPer2.Text = century(mskPer2.Text)
        If IsDate(mskPer2.Text) = False Or Mid(mskPer2.Text, 4, 2) > 12 Then
            MsgBox "Data Inv·lida !", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
        If CDate(mskPer2.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub Option2_Click()

End Sub

Private Sub optTodosRelat_Click()
    cmdGerar.Enabled = False
    If de_informa.rsSel_Relatorios.State = 1 Then de_informa.rsSel_Relatorios.Close
    de_informa.Sel_Relatorios "%"
    gridRelatorios.DataMember = "Sel_relatorios"
End Sub

Private Sub optUsuarioRelat_Click()
    cmdGerar.Enabled = True
    If de_informa.rsSel_Relatorios.State = 1 Then de_informa.rsSel_Relatorios.Close
    de_informa.Sel_Relatorios xusuario
    gridRelatorios.DataMember = "Sel_relatorios"
End Sub

Private Sub txtNumRelat_GotFocus()
    txtNumRelat.SelStart = 0
    txtNumRelat.SelLength = Len(txtNumRelat)
End Sub

Private Sub txtNumRelat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtPara1_GotFocus()
    txtPara1.SelStart = 0
    txtPara1.SelLength = 20
End Sub

Private Sub txtPara1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara1_LostFocus()
    txtPara1 = UCase(txtPara1)
End Sub

Private Sub txtPara2_GotFocus()
    txtPara2.SelStart = 0
    txtPara2.SelLength = 20
End Sub

Private Sub txtPara2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara2_LostFocus()
    txtPara2 = UCase(txtPara2)
End Sub

Private Sub txtPara3_GotFocus()
    txtPara3.SelStart = 0
    txtPara3.SelLength = 20
End Sub

Private Sub txtPara3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara3_LostFocus()
    txtPara3 = UCase(txtPara3)
End Sub

Private Sub txtPara4_GotFocus()
    txtPara4.SelStart = 0
    txtPara4.SelLength = 20
End Sub

Private Sub txtPara4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara4_LostFocus()
    txtPara4 = UCase(txtPara4)
End Sub

Private Sub txtPara5_GotFocus()
    txtPara5.SelStart = 0
    txtPara5.SelLength = 20
End Sub

Private Sub txtPara5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara5_LostFocus()
    txtPara5 = UCase(txtPara5)
End Sub

Private Sub txtPara6_GotFocus()
    txtPara6.SelStart = 0
    txtPara6.SelLength = 20
End Sub

Private Sub txtPara6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara6_LostFocus()
    txtPara6 = UCase(txtPara6)
End Sub

Private Sub txtPara7_GotFocus()
    txtPara7.SelStart = 0
    txtPara7.SelLength = 20
End Sub

Private Sub txtPara7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara7_LostFocus()
    txtPara7 = UCase(txtPara7)
End Sub

Private Sub txtPara8_GotFocus()
    txtPara8.SelStart = 0
    txtPara8.SelLength = 20
End Sub

Private Sub txtPara8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara8_LostFocus()
    txtPara8 = UCase(txtPara8)
End Sub

