VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmConfereRecebCTCs 
   Caption         =   "Conferência de CTCs Recebidos da Emissão"
   ClientHeight    =   6840
   ClientLeft      =   675
   ClientTop       =   1170
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8775
   Begin VB.Frame Frame4 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   8535
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   14
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscaCli 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   8535
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir Não Recebidas"
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir Recebidas"
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   4440
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   7223
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "<ENTER>=Recebido   <ESC>=Não Recebido"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   4440
         Width           =   3270
      End
   End
   Begin VB.CommandButton cmdProcessar 
      Caption         =   "Processar"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   1575
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2985
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmConfereRecebCTCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
