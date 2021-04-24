VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAnOcorr 
   Caption         =   "Análise de Ocorrências"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   435
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRelEmOcorr 
      Caption         =   "Imprimir ""Ocorr Pend"""
      Height          =   315
      Left            =   7560
      TabIndex        =   60
      Top             =   960
      Width           =   1890
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   555
      Left            =   9960
      Picture         =   "frmAnOcorr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   720
      Width           =   690
   End
   Begin VB.CommandButton cmdRelSemPos 
      Caption         =   "Imprimir ""Sem Posição"""
      Height          =   315
      Left            =   7560
      TabIndex        =   55
      Top             =   600
      Width           =   1890
   End
   Begin VB.CommandButton cmdRelEmTransito 
      Caption         =   "Imprimir ""Em Trânsito"""
      Height          =   315
      Left            =   7560
      TabIndex        =   54
      Top             =   240
      Width           =   1890
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   555
      Left            =   10680
      TabIndex        =   24
      Top             =   720
      Width           =   1170
   End
   Begin VB.CommandButton cmdNovaSel 
      Caption         =   "Nova Seleção..."
      Height          =   435
      Left            =   9960
      TabIndex        =   23
      Top             =   240
      Width           =   1890
   End
   Begin VB.Frame Frame4 
      Caption         =   "Gráfico de Ocorrências"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   4920
      TabIndex        =   22
      Top             =   1365
      Width           =   6930
      Begin MSChart20Lib.MSChart GrafOcorrABC 
         Height          =   3270
         Left            =   120
         OleObjectBlob   =   "frmAnOcorr.frx":0496
         TabIndex        =   30
         Top             =   210
         Width           =   6720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ocorrências Mais Frequentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4920
      TabIndex        =   20
      Top             =   5040
      Width           =   6975
      Begin MSDataGridLib.DataGrid GridOcorrABC 
         Bindings        =   "frmAnOcorr.frx":1EFB
         Height          =   1275
         Left            =   120
         TabIndex        =   21
         Top             =   315
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   2249
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8388608
         Enabled         =   -1  'True
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
         DataMember      =   "Sel_ABCOcorr"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "qtdtot"
            Caption         =   "Qtde."
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
            DataField       =   "cod_ocorr"
            Caption         =   "Cod. Ocorr."
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
            DataField       =   "descr_ocorr"
            Caption         =   "Descrição"
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   4935,118
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados Selecionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6825
      Begin VB.Label lblCgcCli 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   480
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4800
         TabIndex        =   57
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Modal:"
         Height          =   195
         Left            =   4200
         TabIndex        =   56
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblDataPer2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3045
         TabIndex        =   17
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2835
         TabIndex        =   16
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblDataPer1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   4785
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Período..............: De"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cliente / Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   120
      TabIndex        =   0
      Top             =   1365
      Width           =   4770
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         TabIndex        =   52
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblPercInf7b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4080
         TabIndex        =   48
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label lblPercInf6b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4080
         TabIndex        =   47
         Top             =   4320
         Width           =   540
      End
      Begin VB.Label lblPercInf5b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   45
         Top             =   3840
         Width           =   540
      End
      Begin VB.Label lblPercInf4b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   46
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label lblPercInf7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2640
         TabIndex        =   31
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label lblPercInf6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2640
         TabIndex        =   32
         Top             =   4320
         Width           =   540
      End
      Begin VB.Label lblPercInf5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   35
         Top             =   3840
         Width           =   540
      End
      Begin VB.Label lblPercInf4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   33
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label lblInf2 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2880
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblInf1 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblInf7b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3360
         TabIndex        =   50
         Top             =   4920
         Width           =   750
      End
      Begin VB.Label lblInf6b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3360
         TabIndex        =   49
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label lblInf5b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Top             =   3840
         Width           =   750
      End
      Begin VB.Label lblInf4totb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   42
         Top             =   3360
         Width           =   750
      End
      Begin VB.Label lblInf7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   4920
         Width           =   750
      End
      Begin VB.Label lblInf6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label lblInf5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   36
         Top             =   3840
         Width           =   750
      End
      Begin VB.Label lblInf4tot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   3360
         Width           =   750
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000003&
         X1              =   3240
         X2              =   3240
         Y1              =   2160
         Y2              =   5280
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         X1              =   4680
         X2              =   4680
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   120
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblInf41b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label lblInf42b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3360
         TabIndex        =   43
         Top             =   3000
         Width           =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   3240
         X2              =   3240
         Y1              =   1200
         Y2              =   2160
      End
      Begin VB.Label lblPercInf3b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   40
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label lblInf3b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   41
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "NFs"
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
         Left            =   3840
         TabIndex        =   39
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CTCs"
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
         Left            =   2280
         TabIndex        =   38
         Top             =   1320
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblInf42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   3000
         Width           =   750
      End
      Begin VB.Label lblInf41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "5 - Em Trânsito ...........:"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   3840
         Width           =   1650
      End
      Begin VB.Label lblPercInf3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   34
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Ocorrências..:"
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "4.2 - Pendentes.....:"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "4.1 - Baixados.......:"
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label19 
         Caption         =   "7 - Total Pendentes              (Ítens: 4.2 + 6) ........:"
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   1830
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "6 - Sem Posição..........:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   1650
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "4 - NÃO Entregues e com Ocorrências:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   2760
      End
      Begin VB.Label lblInf3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "3 - Entregues .............:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "2 - Qtde. total de NFs..........:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1 - Qtde. total de CTCs .......:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmAnOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdImpr_Click()

End Sub

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdNovaSel_Click()
    Unload frmAnOcorr
    frmEscCliPer.Caption = "Análise de Ocorrências"
    frmEscCliPer.Show 1
End Sub

Private Sub cmdRelEmTransito_Click()
Dim xmodal As String
    If lblModal = "RODO/AÉREO" Then
        xmodal = "%"
    Else
        xmodal = Mid(lblModal, 1, 1) & "%"
    End If
    If de_informa.rsSel_CtcEmTransitoREL.State = 1 Then de_informa.rsSel_CtcEmTransitoREL.Close
    de_informa.Sel_CtcEmTransitoREL CDate(lblDataPer1), CDate(lblDataPer2), lblCgcCli, "N", datahora("data"), xmodal
    drEmTransito.Show 1
End Sub

Private Sub cmdRelResumo_Click()

End Sub

Private Sub cmdRelSemPos_Click()
Dim xmodal As String
    If lblModal = "RODO/AÉREO" Then
        xmodal = "%"
    Else
        xmodal = Mid(lblModal, 1, 1) & "%"
    End If
    If de_informa.rsSel_CtcPendenteEntrREL.State = 1 Then de_informa.rsSel_CtcPendenteEntrREL.Close
    de_informa.Sel_CtcPendenteEntrREL CDate(lblDataPer1), CDate(lblDataPer2), lblCgcCli, "N", datahora("data"), xmodal
    drPendente.Show 1
End Sub

Private Sub cmdSair_Click()
    Unload frmEscCliPer
    Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
'    Unload frmEscCliPer
    'cmdNovaSel.SetFocus
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmAnOcorr = Nothing
End Sub

