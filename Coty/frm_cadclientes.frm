VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cadclientes 
   Caption         =   "CADASTRO DE CLIENTES"
   ClientHeight    =   8220
   ClientLeft      =   1905
   ClientTop       =   1320
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txt_obs 
         Height          =   1575
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   4920
         Width           =   6135
      End
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   255
         Left            =   10200
         TabIndex        =   30
         Top             =   7560
         Width           =   1335
      End
      Begin VB.TextBox txt_uf 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8280
         TabIndex        =   29
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   4800
         TabIndex        =   27
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txt_complemento 
         Height          =   285
         Left            =   10320
         TabIndex        =   25
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txt_numero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8280
         TabIndex        =   23
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txt_email 
         Height          =   285
         Left            =   9000
         TabIndex        =   9
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox txt_contato 
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Top             =   3960
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mask_cep 
         Height          =   300
         Left            =   3480
         TabIndex        =   7
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_fone 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txt_endereco 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   2880
         Width           =   6135
      End
      Begin VB.TextBox txt_cliente_fantasia 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txt_cgc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9000
         TabIndex        =   2
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Label Label17 
         Caption         =   "OBS:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "UF:"
         Height          =   255
         Left            =   7920
         TabIndex        =   28
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "CIDADE:"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "COMPLEMENTO:"
         Height          =   255
         Left            =   9000
         TabIndex        =   24
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Nº:"
         Height          =   255
         Left            =   7920
         TabIndex        =   22
         Top             =   2880
         Width           =   255
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   11640
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label12 
         Caption         =   "E-MAIL:"
         Height          =   255
         Left            =   7920
         TabIndex        =   21
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "NOME CONTATO:"
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "CEP:"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "TELEFONE:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "BAIRRO:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "ENDEREÇO:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "NOME FANTASIA:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "CNPJ / CGC:"
         Height          =   255
         Left            =   7920
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "NOME DO CLIENTE:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "CADASTRO DE CLIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   11640
         X2              =   120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Mensageiros Motorizados S/C Ltda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "COTY MOTOS"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   480
         Width           =   7215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   9600
         Picture         =   "frm_cadclientes.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_cadclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_sair_Click()

MDIForm1.Toolbar1.Enabled = True

Unload Me


End Sub
