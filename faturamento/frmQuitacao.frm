VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmQuitacao 
   Caption         =   "Quitar Fatura"
   ClientHeight    =   4665
   ClientLeft      =   1770
   ClientTop       =   1455
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdGravarDesconto 
      Caption         =   "Quitar Fatura"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quitação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtObsPag 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   3
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txtObsAcresc 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtAcrescimo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   3015
         Begin VB.Label lblFilialFatura 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial-Fatura:"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   840
         End
      End
      Begin MSMask.MaskEdBox mskPagto 
         Height          =   285
         Left            =   2040
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "( = )  Valor Recebido:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label lblValorRecebido 
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
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "( + )  Acréscimos:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "( = )  Valor da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label lblValorFatura 
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
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Pagamento:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmQuitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravarDesconto_Click()

    If Not IsDate(mskPagto) Then
        MsgBox "Data Inválida para Pagamento desta Fatura !"
        mskPagto.SetFocus
        Exit Sub
    End If
    
    If CDbl(SoNumeros(txtAcrescimo)) / 100 > 0 And Len(Trim$(txtObsAcresc)) < 3 Then
        MsgBox "Informe o Motivo deste Acréscimo !"
        txtObsAcresc.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtAcrescimo)) = 0 Then
        txtAcrescimo = "0"
    End If
    
     If MsgBox("Você Confirma a QUITAÇÃO Desta Fatura ?", vbYesNo + vbQuestion, "Desconto") = vbYes Then
        de_informa.Alt_QuitacaoFatura CDate(mskPagto), xusuario, CDbl(SoNumeros(txtAcrescimo)) / 100, txtObsAcresc, xusuario, CDbl(SoNumeros(txtAcrescimo)) / 100, txtObsPag, lblFilialFatura
        Unload Me
    End If
   
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub mskPagto_GotFocus()
    mskPagto.SelStart = 0
    mskPagto.SelLength = 10
End Sub
Private Sub mskPagto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(mskPagto)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub
Private Sub mskPagto_LostFocus()
    If mskPagto.Text <> "__/__/____" Then
        mskPagto.Text = century(mskPagto.Text)
        If IsDate(mskPagto.Text) = False Or Mid(mskPagto.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPagto.SetFocus
            Exit Sub
        End If
        If CDate(mskPagto.Text) < CDate(lblEmissao.Caption) Then
            MsgBox "ATENÇÃO ! A Data de Pagamento não pode ser Inferior a Data de Emissão !!!", vbCritical, "Erro"
            mskPagto.SetFocus
        End If
        If CDate(mskPagto.Text) > datahora("DATA") Then
            MsgBox "ATENÇÃO ! A Data de Pagamento não pode ser Maior que Hoje !!!", vbCritical, "Erro"
            mskPagto.SetFocus
        End If
    End If
End Sub
Private Sub txtAcrescimo_Change()
    If Not IsNumeric(txtAcrescimo) Then
        SendKeys "{BACKSPACE}"
        Exit Sub
    End If
    Call TextMoneyBox_Change(txtAcrescimo)
    DoEvents
    If Len(Trim$(txtAcrescimo)) > 0 Then
        If CDbl(SoNumeros(txtAcrescimo)) / 100 > 0 Then
            txtObsAcresc.Enabled = True
            txtObsAcresc.BackColor = xamarelo1
            DoEvents
            Exit Sub
        End If
    End If
    txtObsAcresc.Enabled = False
    txtObsAcresc.BackColor = xbranco
    DoEvents
End Sub

Private Sub txtAcrescimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAcrescimo_LostFocus()
    If Len(Trim$(txtAcrescimo)) > 0 Then
        If CDbl(SoNumeros(txtAcrescimo)) / 100 > 0 Then
            If Len(Trim$(txtAbat)) = 0 Then txtAbat = 0
            lblValorRecebido = Format((CDbl(SoNumeros(lblValorFatura)) / 100) + (CDbl(SoNumeros(txtAcrescimo)) / 100), "##,###,##0.00")
            txtObsAcresc.Enabled = True
            txtObsAcresc.BackColor = xamarelo1
            DoEvents
            txtObsAcresc.SetFocus
            Exit Sub
        End If
    End If
    txtObsAcresc.Enabled = False
    txtObsAcresc.BackColor = xbranco
    DoEvents
End Sub
Private Sub txtObsPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtObsPag_LostFocus()
    txtObsPag = UCase(Trim$(txtObsPag))
End Sub


