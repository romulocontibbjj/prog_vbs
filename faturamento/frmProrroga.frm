VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmProrroga 
   Caption         =   "Prorrogar Vencimento"
   ClientHeight    =   4395
   ClientLeft      =   2370
   ClientTop       =   1605
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6030
   Begin VB.Frame Frame1 
      Caption         =   "Prorrogação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5775
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   2535
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial-Fatura:"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   840
         End
         Begin VB.Label lblFilialFatura 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtObsProrroga 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   1
         Top             =   2280
         Width           =   5175
      End
      Begin MSMask.MaskEdBox mskDataProrr 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "* Na Obs Informe o Nome de Quem Autorizou a Prorrogação *"
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
         TabIndex        =   14
         Top             =   2760
         Width           =   5595
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Prorrogar Vencimento Para:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1950
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdGravarDesconto 
      Caption         =   "Prorrogar Fatura"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3720
      Width           =   2055
   End
End
Attribute VB_Name = "frmProrroga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravarDesconto_Click()

    If Not IsDate(mskDataProrr) Then
        MsgBox "Data Para Prorrogação Inválida !"
        mskDataProrr.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtObsProrroga)) < 3 Then
        MsgBox "Por Favor Informe na Observação Quem Autorizou a Prorrogação Desta Fatura !", vbInformation, "Observação"
        txtObsProrroga.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Você Confirma a Prorrogação de Vencimento desta Fatura ?", vbYesNo + vbQuestion, "Prorrogação") = vbYes Then
        de_informa.Alt_ProrrogaFatura CDate(mskDataProrr), xusuario, txtObsProrroga, lblFilialFatura
        de_informa.Ins_FaturaHistorico lblFilialFatura, xusuario, "PRORROGACAO", txtObsProrroga, ""
        Unload Me
    End If

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub mskDataProrr_GotFocus()
    mskDataProrr.SelStart = 0
    mskDataProrr.SelLength = 10
End Sub
Private Sub mskDataProrr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(mskDataProrr)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskDataProrr_LostFocus()
    If mskDataProrr.Text <> "__/__/____" Then
        mskDataProrr.Text = century(mskDataProrr.Text)
        If IsDate(mskDataProrr.Text) = False Or Mid(mskDataProrr.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskDataProrr.SetFocus
            Exit Sub
        End If
        If CDate(mskDataProrr.Text) < CDate(lblEmissao.Caption) Then
            MsgBox "ATENÇÃO ! A Data de Vencimento não pode ser Inferior a Data de Emissão !!!", vbCritical, "Erro"
            mskDataProrr.SetFocus
        End If
    End If
End Sub
Private Sub txtObsProrroga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtObsProrroga_LostFocus()
    txtObsProrroga = UCase(Trim$(txtObsProrroga))
End Sub



