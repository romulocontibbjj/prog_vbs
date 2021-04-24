VERSION 5.00
Begin VB.Form frmEDIConf 
   Caption         =   "Confirmação do Lancamento de Arquivo EDI"
   ClientHeight    =   2505
   ClientLeft      =   2685
   ClientTop       =   1530
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   6375
   Begin VB.Frame Frame1 
      Caption         =   "Confirmação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdCanc 
         BackColor       =   &H80000004&
         Caption         =   "Cancelar Processo Restante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdConfirma 
         BackColor       =   &H80000004&
         Caption         =   "Confirma Lançamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdIgnora 
         BackColor       =   &H80000004&
         Caption         =   "Ignora Lançamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Você confirma o lançamento do registro do arquivo EDI no Sistema Intec ?"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   5280
      End
      Begin VB.Label Label1 
         Caption         =   "Confira atentamente os dados contidos atualmente no Sistema com os dados contidos no arquivo EDI."
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmEDIConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCanc_Click()
    frmEdiImport.lblConfirma.Caption = "CANCELAR"
    Me.Hide
End Sub

Private Sub cmdConfirma_Click()
    frmEdiImport.lblConfirma.Caption = "SIM"
    Me.Hide
End Sub

Private Sub cmdIgnora_Click()
    frmEdiImport.lblConfirma.Caption = "NÃO"
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Top = 1000
    Me.Left = 2700
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEDIConf = Nothing
End Sub
