VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadPrazos 
   Caption         =   "Cadastro de Tabelas de Prazos de Entrega"
   ClientHeight    =   6120
   ClientLeft      =   -30
   ClientTop       =   510
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdImprTela 
      Height          =   495
      Left            =   10440
      Picture         =   "frmCadPrazos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   720
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid gridCadPrazo 
      Bindings        =   "frmCadPrazos.frx":0772
      Height          =   1095
      Left            =   1560
      TabIndex        =   76
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
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
      DataMember      =   "Sel_TabPrazoGro"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "codigo"
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
            ColumnWidth     =   900,284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8880
      TabIndex        =   75
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Canc. / Sair"
      Height          =   495
      Left            =   11160
      TabIndex        =   74
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdAltera 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      TabIndex        =   73
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdNovaTab 
      Caption         =   "Nova Tabela"
      Height          =   375
      Left            =   8880
      TabIndex        =   72
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame fraAereo 
      Caption         =   "AÉREO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   6120
      TabIndex        =   37
      Top             =   1320
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid FlexAir1 
         Height          =   4095
         Left            =   1440
         TabIndex        =   38
         Top             =   480
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   17
         FixedCols       =   0
         ScrollBars      =   0
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid FlexAir2 
         Height          =   2895
         Left            =   4320
         TabIndex        =   39
         Top             =   480
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   12
         FixedCols       =   0
         ScrollBars      =   0
         Appearance      =   0
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000010&
         X1              =   1440
         X2              =   120
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000010&
         X1              =   960
         X2              =   120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   1440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label64 
         Caption         =   "REGIÃO NORTE"
         Height          =   435
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "REGIÃO NORDESTE"
         Height          =   435
         Left            =   120
         TabIndex        =   70
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label62 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AC"
         Height          =   255
         Left            =   1080
         TabIndex        =   69
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label57 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AM"
         Height          =   255
         Left            =   1080
         TabIndex        =   68
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label56 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AP"
         Height          =   255
         Left            =   1080
         TabIndex        =   67
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label55 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PA"
         Height          =   255
         Left            =   1080
         TabIndex        =   66
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label54 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RO"
         Height          =   255
         Left            =   1080
         TabIndex        =   65
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RR"
         Height          =   255
         Left            =   1080
         TabIndex        =   64
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label52 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TO"
         Height          =   255
         Left            =   1080
         TabIndex        =   63
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label51 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AL"
         Height          =   255
         Left            =   1080
         TabIndex        =   62
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label50 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BA"
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label49 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SE"
         Height          =   255
         Left            =   1080
         TabIndex        =   60
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label48 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PE"
         Height          =   255
         Left            =   1080
         TabIndex        =   59
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label47 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PB"
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label46 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RN"
         Height          =   255
         Left            =   1080
         TabIndex        =   57
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CE"
         Height          =   255
         Left            =   1080
         TabIndex        =   56
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PI"
         Height          =   255
         Left            =   1080
         TabIndex        =   55
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label43 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MA"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   4320
         Width           =   375
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000010&
         X1              =   3840
         X2              =   3000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label42 
         Caption         =   "REGIÃO SUL"
         Height          =   435
         Left            =   3000
         TabIndex        =   53
         Top             =   1800
         Width           =   855
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000010&
         X1              =   4320
         X2              =   3000
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   3840
         X2              =   3000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000010&
         X1              =   3000
         X2              =   4320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label41 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MT"
         Height          =   255
         Left            =   3960
         TabIndex        =   52
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MS"
         Height          =   255
         Left            =   3960
         TabIndex        =   51
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label39 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GO"
         Height          =   255
         Left            =   3960
         TabIndex        =   50
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label38 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DF"
         Height          =   255
         Left            =   3960
         TabIndex        =   49
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label37 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SC"
         Height          =   255
         Left            =   3960
         TabIndex        =   48
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label36 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RS"
         Height          =   255
         Left            =   3960
         TabIndex        =   47
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PR"
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SP"
         Height          =   255
         Left            =   3960
         TabIndex        =   45
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RJ"
         Height          =   255
         Left            =   3960
         TabIndex        =   44
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MG"
         Height          =   255
         Left            =   3960
         TabIndex        =   43
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ES"
         Height          =   255
         Left            =   3960
         TabIndex        =   42
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "REGIÃO C.OESTE"
         Height          =   435
         Left            =   3000
         TabIndex        =   41
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "REGIÃO SUDESTE"
         Height          =   435
         Left            =   3000
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame fraCodTabela 
      Caption         =   "Cód. Tabela"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtCodTabela 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame fraRodo 
      Caption         =   "RODOVIÁRIO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid FlexRodo1 
         Height          =   4095
         Left            =   1440
         TabIndex        =   35
         Top             =   480
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   17
         FixedCols       =   0
         ScrollBars      =   0
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid FlexRodo2 
         Height          =   2895
         Left            =   4320
         TabIndex        =   36
         Top             =   480
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   12
         FixedCols       =   0
         ScrollBars      =   0
         Appearance      =   0
      End
      Begin VB.Label Label21 
         Caption         =   "REGIÃO SUDESTE"
         Height          =   435
         Left            =   3000
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "REGIÃO C.OESTE"
         Height          =   435
         Left            =   3000
         TabIndex        =   31
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ES"
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MG"
         Height          =   255
         Left            =   3960
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RJ"
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SP"
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PR"
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RS"
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SC"
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label58 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DF"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label59 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GO"
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label60 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MS"
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label61 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MT"
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   3120
         Width           =   375
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   3000
         X2              =   4320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   3840
         X2              =   3000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000010&
         X1              =   4320
         X2              =   3000
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label31 
         Caption         =   "REGIÃO SUL"
         Height          =   435
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   3840
         X2              =   3000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MA"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PI"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CE"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RN"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PB"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PE"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SE"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BA"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AL"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TO"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RR"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RO"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PA"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AP"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AM"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AC"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "REGIÃO NORDESTE"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "REGIÃO NORTE"
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   1440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   960
         X2              =   120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   1440
         X2              =   120
         Y1              =   4560
         Y2              =   4560
      End
   End
   Begin MSDataGridLib.DataGrid gridClientes 
      Bindings        =   "frmCadPrazos.frx":078B
      Height          =   1095
      Left            =   3120
      TabIndex        =   77
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1931
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
      DataMember      =   "Sel_ClienteTabPrz"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "nome"
         Caption         =   "Clientes que Utilizam esta Tabela de Prazos"
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
            ColumnWidth     =   5040
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadPrazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Click()
    
End Sub

Private Sub cmdAltera_Click()
    'fraRodo.Enabled = True
    'fraAereo.Enabled = True
End Sub

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdNovaTab_Click()
    If Mid$(xdireitos, 9, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmNovaTabPrz.Show 1
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub FlexRodo1_Click()
    MsgBox FlexRodo1.Row
    
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    FlexRodo1.ColWidth(0) = 653
    FlexRodo1.ColWidth(1) = 653
    FlexRodo2.ColWidth(0) = 653
    FlexRodo2.ColWidth(1) = 653
    FlexAir1.ColWidth(0) = 653
    FlexAir1.ColWidth(1) = 653
    FlexAir2.ColWidth(0) = 653
    FlexAir2.ColWidth(1) = 653
    FlexRodo1.Row = 0
    FlexRodo1.Col = 0
    FlexRodo1.Text = "Capital"
    FlexRodo1.Col = 1
    FlexRodo1.Text = "Interior"
    FlexRodo2.Row = 0
    FlexRodo2.Col = 0
    FlexRodo2.Text = "Capital"
    FlexRodo2.Col = 1
    FlexRodo2.Text = "Interior"
    FlexAir1.Row = 0
    FlexAir1.Col = 0
    FlexAir1.Text = "Capital"
    FlexAir1.Col = 1
    FlexAir1.Text = "Interior"
    FlexAir2.Row = 0
    FlexAir2.Col = 0
    FlexAir2.Text = "Capital"
    FlexAir2.Col = 1
    FlexAir2.Text = "Interior"
    gridCadPrazo.DataMember = "sel_tabprazogro"
    gridCadPrazo.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmCadPrazos = Nothing
End Sub

Private Sub gridCadPrazo_Click()
    txtCodTabela = gridCadPrazo.Columns(0)
    Call atualiza_rodo
    Call atualiza_air
    If de_informa.rsSel_ClienteTabPrz.State = 1 Then de_informa.rsSel_ClienteTabPrz.Close
    de_informa.Sel_ClienteTabPrz gridCadPrazo.Columns(0)
    gridClientes.DataMember = "sel_clientetabprz"
    gridClientes.Refresh
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "CONSULTA", xusuario, "CAD. DE TABELAS DE PRAZOS DE ENTREGA: " & txtCodTabela
    
    'cmdAltera.Enabled = True
End Sub
Private Sub atualiza_rodo()
    If de_informa.rsSel_CadPrazo.State = 1 Then de_informa.rsSel_CadPrazo.Close
    de_informa.Sel_CadPrazo gridCadPrazo.Columns(0), "R"
    de_informa.rsSel_CadPrazo.MoveFirst
    Do Until de_informa.rsSel_CadPrazo.EOF
        If de_informa.rsSel_CadPrazo.Fields("uf") = "AC" Then
            FlexRodo1.Row = 1
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AM" Then
            FlexRodo1.Row = 2
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AP" Then
            FlexRodo1.Row = 3
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PA" Then
            FlexRodo1.Row = 4
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RO" Then
            FlexRodo1.Row = 5
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RR" Then
            FlexRodo1.Row = 6
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "TO" Then
            FlexRodo1.Row = 7
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AL" Then
            FlexRodo1.Row = 8
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "BA" Then
            FlexRodo1.Row = 9
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SE" Then
            FlexRodo1.Row = 10
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PE" Then
            FlexRodo1.Row = 11
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PB" Then
            FlexRodo1.Row = 12
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RN" Then
            FlexRodo1.Row = 13
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "CE" Then
            FlexRodo1.Row = 14
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PI" Then
            FlexRodo1.Row = 15
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MA" Then
            FlexRodo1.Row = 16
            FlexRodo1.Col = 0
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo1.Col = 1
            FlexRodo1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "ES" Then
            FlexRodo2.Row = 1
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MG" Then
            FlexRodo2.Row = 2
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RJ" Then
            FlexRodo2.Row = 3
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SP" Then
            FlexRodo2.Row = 4
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PR" Then
            FlexRodo2.Row = 5
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RS" Then
            FlexRodo2.Row = 6
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SC" Then
            FlexRodo2.Row = 7
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "DF" Then
            FlexRodo2.Row = 8
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "GO" Then
            FlexRodo2.Row = 9
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MS" Then
            FlexRodo2.Row = 10
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MT" Then
            FlexRodo2.Row = 11
            FlexRodo2.Col = 0
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexRodo2.Col = 1
            FlexRodo2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        End If
        de_informa.rsSel_CadPrazo.MoveNext
    Loop
End Sub

Private Sub atualiza_air()
    If de_informa.rsSel_CadPrazo.State = 1 Then de_informa.rsSel_CadPrazo.Close
    de_informa.Sel_CadPrazo gridCadPrazo.Columns(0), "A"
    de_informa.rsSel_CadPrazo.MoveFirst
    Do Until de_informa.rsSel_CadPrazo.EOF
        If de_informa.rsSel_CadPrazo.Fields("uf") = "AC" Then
            FlexAir1.Row = 1
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AM" Then
            FlexAir1.Row = 2
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AP" Then
            FlexAir1.Row = 3
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PA" Then
            FlexAir1.Row = 4
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RO" Then
            FlexAir1.Row = 5
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RR" Then
            FlexAir1.Row = 6
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "TO" Then
            FlexAir1.Row = 7
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AL" Then
            FlexAir1.Row = 8
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "BA" Then
            FlexAir1.Row = 9
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SE" Then
            FlexAir1.Row = 10
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PE" Then
            FlexAir1.Row = 11
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PB" Then
            FlexAir1.Row = 12
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RN" Then
            FlexAir1.Row = 13
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "CE" Then
            FlexAir1.Row = 14
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PI" Then
            FlexAir1.Row = 15
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MA" Then
            FlexAir1.Row = 16
            FlexAir1.Col = 0
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir1.Col = 1
            FlexAir1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "ES" Then
            FlexAir2.Row = 1
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MG" Then
            FlexAir2.Row = 2
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RJ" Then
            FlexAir2.Row = 3
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SP" Then
            FlexAir2.Row = 4
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PR" Then
            FlexAir2.Row = 5
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RS" Then
            FlexAir2.Row = 6
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SC" Then
            FlexAir2.Row = 7
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "DF" Then
            FlexAir2.Row = 8
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "GO" Then
            FlexAir2.Row = 9
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MS" Then
            FlexAir2.Row = 10
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MT" Then
            FlexAir2.Row = 11
            FlexAir2.Col = 0
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
            FlexAir2.Col = 1
            FlexAir2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
        End If
        de_informa.rsSel_CadPrazo.MoveNext
    Loop
End Sub

