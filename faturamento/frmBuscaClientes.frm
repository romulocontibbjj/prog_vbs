VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaSubClientes 
   Caption         =   "Busca Cadastro de SubContratados - Cadastro"
   ClientHeight    =   5040
   ClientLeft      =   675
   ClientTop       =   1125
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9165
   Begin VB.Frame fraConsCli 
      Caption         =   "Busca..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8925
      Begin VB.CommandButton cmdCancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Confirma"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   4200
         Width           =   1335
      End
      Begin VB.OptionButton optBuscaCGC 
         Caption         =   "Busca por CGC"
         Height          =   195
         Left            =   5040
         TabIndex        =   6
         Top             =   3360
         Width           =   1575
      End
      Begin VB.OptionButton optBuscaFantasia 
         Caption         =   "Busca por Nome Fantasia"
         Height          =   195
         Left            =   5040
         TabIndex        =   7
         Top             =   3600
         Width           =   2175
      End
      Begin VB.OptionButton optBuscaRazaoTodo 
         Caption         =   "Busca no Texto Todo..."
         Height          =   195
         Left            =   5040
         TabIndex        =   5
         Top             =   3150
         Width           =   2220
      End
      Begin VB.OptionButton optBuscaRazapInic 
         Caption         =   "Busca no Início do Texto..."
         Height          =   195
         Left            =   5040
         TabIndex        =   4
         Top             =   2940
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   3120
         Width           =   825
      End
      Begin VB.TextBox txtBuscaSubCli 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   0
         Top             =   3120
         Width           =   2265
      End
      Begin MSDataGridLib.DataGrid gridBuscaSubCli 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   -2147483634
         ForeColor       =   -2147483630
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
         DataMember      =   "Sel_CadCliCGCLike"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "cgc"
            Caption         =   "CNPJ / CPF"
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
            Caption         =   "Nome / Razão Soc."
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
            DataField       =   "fantasia"
            Caption         =   "Nome Fantasia"
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
            DataField       =   "apelido"
            Caption         =   "Apelido"
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
            DataField       =   "cidade"
            Caption         =   "Cidade"
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
            DataField       =   "uf"
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
         BeginProperty Column06 
            DataField       =   "rem_des_log"
            Caption         =   "rem_des_log"
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
            DataField       =   "endereco"
            Caption         =   "Endereço"
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
            DataField       =   "complemento"
            Caption         =   "complemento"
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
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3075,024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2039,811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   374,74
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social"
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
         Left            =   7440
         TabIndex        =   11
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Line Line1 
         X1              =   7290
         X2              =   7290
         Y1              =   2955
         Y2              =   3325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Nome:"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   3120
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmBuscaSubClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gridBuscaSubCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
'        cmdConfirma_Click
    End If
    If KeyAscii = 27 Then   'TECLA ENTER
'        cmdCancela_Click
    End If
End Sub

Private Sub txtBuscaSubCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
    If KeyAscii = 27 Then   'TECLA ENTER
'        cmdCancela_Click
    End If
End Sub
