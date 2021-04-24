VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPesquisaFatura 
   Caption         =   "Pesquisar Fatura"
   ClientHeight    =   5790
   ClientLeft      =   720
   ClientTop       =   900
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9015
   Begin VB.Frame fraDados 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "Confirmar "
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   3480
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid gridRelFatura 
         Bindings        =   "frmPesquisaFatura.frx":0000
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5530
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
         DataMember      =   "Sel_RelFaturas1"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Fatura"
            Caption         =   "Filial-Fatura"
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
            DataField       =   "emissao"
            Caption         =   "Emissão"
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
            DataField       =   "cliente_nome"
            Caption         =   "Cliente Nome"
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
            DataField       =   "vencimento"
            Caption         =   "Vencimento"
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
            DataField       =   "valorfatura"
            Caption         =   "Valor Fatura"
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
            DataField       =   "obsfatura"
            Caption         =   "Observação"
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
            DataField       =   "status"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2910,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2954,835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   299,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "No Período de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7305
      Begin VB.CommandButton cmdBuscaCli 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   14
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Buscar ..."
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
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
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   12
         Top             =   960
         Width           =   4200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmPesquisaFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
