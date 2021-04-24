VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmAlteraVencPrefat 
   Caption         =   "Altera Vencimento Pré-Fatura"
   ClientHeight    =   3045
   ClientLeft      =   1500
   ClientTop       =   1560
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5775
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Gravar Novo Vencimento"
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   2160
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mskVencimento 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblUsu 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblVencAtual 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento Atual:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label lblPrefat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pré-Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Novo Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmAlteraVencPrefat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        
    If IsDate(mskVencimento.Text) = False Or Mid(mskVencimento.Text, 4, 2) > 12 Then
        MsgBox "Data Inválida !", vbCritical, "Erro"
        mskVencimento.SetFocus
        Exit Sub
    End If
    If CDate(mskVencimento.Text) < datahora("data") Then
        MsgBox "ERRO ! Data de Vencimento Menor que Hoje ???", vbCritical, "Erro"
        mskVencimento.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Você Confirma esta Nova Data de Vencimento: " & mskVencimento & " ?", vbYesNo + vbQuestion, "Confirmação") = vbYes Then
        de_informa.Alt_VencPrefatura CDate(mskVencimento), lblPrefat
        MsgBox "OK ! Data de Vencimento Alterada !", vbInformation, "OK"
    Else
        MsgBox "Data de Vencimento NÃO Alterada !", vbCritical, "OPS"
    End If
    Unload Me
    
End Sub

Private Sub mskVencimento_GotFocus()
    mskVencimento.SelStart = 0
    mskVencimento.SelLength = 10
End Sub

Private Sub mskVencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(mskVencimento)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskVencimento_LostFocus()
    If mskVencimento.Text <> "__/__/____" Then
        mskVencimento.Text = century(mskVencimento.Text)
        If IsDate(mskVencimento.Text) = False Or Mid(mskVencimento.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskVencimento.SetFocus
            Exit Sub
        End If
        If CDate(mskVencimento.Text) < datahora("data") Then
            MsgBox "ERRO ! Data de Vencimento Menor que Hoje ???", vbCritical, "Erro"
            mskVencimento.SetFocus
            Exit Sub
        End If
        Command1.Enabled = True
    End If
End Sub

