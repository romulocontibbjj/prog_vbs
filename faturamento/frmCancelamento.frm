VERSION 5.00
Begin VB.Form frmCancelamento 
   Caption         =   "Cancelamento de Fatura"
   ClientHeight    =   6525
   ClientLeft      =   1080
   ClientTop       =   990
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7215
   Begin VB.Frame Frame1 
      Caption         =   "Cancelamento"
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
      TabIndex        =   36
      Top             =   5160
      Width           =   6975
      Begin VB.TextBox txtObsCanc 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   300
         Width           =   6735
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "S A I R"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C A N C E L A R"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Número da Fatura"
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
      TabIndex        =   30
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFatura 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Filial Fatura:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Dados da Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6975
      Begin VB.Frame Frame3 
         Caption         =   "Status"
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
         Left            =   3240
         TabIndex        =   32
         Top             =   2040
         Width           =   3615
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Caption         =   "STATUS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   35
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Label lblAcrescimo 
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
         Left            =   1800
         TabIndex        =   34
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblAbat 
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
         Left            =   1800
         TabIndex        =   33
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencto:"
         Height          =   195
         Left            =   2400
         TabIndex        =   27
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblBanco 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lblConta 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5400
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   4680
         TabIndex        =   23
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblValorFaturaBruto 
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
         Left            =   1800
         TabIndex        =   21
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "( - )  Abatimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "( + )  Acréscimos:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "( = )  Valor da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3720
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
         Left            =   1800
         TabIndex        =   16
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblClienteCNPJ 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Emissor:"
         Height          =   195
         Left            =   4680
         TabIndex        =   12
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblEmissor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5400
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto com ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblValorFaturaBrutoICMS 
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
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "( - ) Desc. ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1980
         Width           =   1170
      End
      Begin VB.Label lblValorICMS 
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
         Left            =   1800
         TabIndex        =   7
         Top             =   1980
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCancelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    If MsgBox("Você Confirma o Cancelamento desta Fatura ?", vbYesNo + vbQuestion, "Confirma") = vbYes Then
        de_informa.cn_informa.BeginTrans
        
            'cancela a fatura
            de_informa.Alt_CancelarFatura xusuario, Trim$(txtObsCanc), TransFatur(txtFilial, txtFatura)
            'deleta os itens
            de_informa.Excl_CancFaturaItens TransFatur(txtFilial, txtFatura)
            'limpa os CTCs
            de_informa.Alt_CancFatLimpaCTC TransFatur(txtFilial, txtFatura)
            'limpa as NFS
            de_informa.Alt_CancFatLimpaNFS TransFatur(txtFilial, txtFatura)
        
        de_informa.cn_informa.CommitTrans
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim xFilialFatura As String
    
    xFilialFatura = TransFatur(txtFilial, txtFatura)
    
    If de_informa.rsSel_Fatura.State = 1 Then de_informa.rsSel_Fatura.Close
    de_informa.Sel_Fatura xFilialFatura
    
    If de_informa.rsSel_Fatura.RecordCount > 0 Then
    
        lblValorFaturaBrutoICMS = de_informa.rsSel_Fatura.Fields("valorbrutoicms")
        lblValorICMS = de_informa.rsSel_Fatura.Fields("descicms")
        lblValorFaturaBruto = de_informa.rsSel_Fatura.Fields("valorbruto")
        txtAbat = de_informa.rsSel_Fatura.Fields("abatimento")
        lblTipoAbat = de_informa.rsSel_Fatura.Fields("tipoabat")
        txtObsAbat = de_informa.rsSel_Fatura.Fields("obsabat")
        lblValorFatura = de_informa.rsSel_Fatura.Fields("valorfatura")
        txtObsFatura = de_informa.rsSel_Fatura.Fields("obsfatura")
        
        lblClienteCNPJ = de_informa.rsSel_Fatura.Fields("cliente_cgc")
        lblCliente = de_informa.rsSel_Fatura.Fields("cliente_nome")
        lblEmissao = de_informa.rsSel_Fatura.Fields("emissao")
        lblVencto = de_informa.rsSel_Fatura.Fields("vencimento")
        lblEmissor = de_informa.rsSel_Fatura.Fields("emissor")
        lblBanco = de_informa.rsSel_Fatura.Fields("banconome")
        lblConta = de_informa.rsSel_Fatura.Fields("conta")
        If de_informa.rsSel_Fatura.Fields("status") = "C" Then
            lblStatus = "CANCELADO"
            cmdCancelar.Enabled = False
        ElseIf de_informa.rsSel_Fatura.Fields("status") = "Q" Then
            lblStatus = "QUITADO"
            cmdCancelar.Enabled = False
        ElseIf de_informa.rsSel_Fatura.Fields("status") = "N" Then
            lblStatus = "EM ABERTO"
            cmdCancelar.Enabled = True
        End If
    
    Else
    
        MsgBox "Fatura Não Encontrada !", vbInformation, "Erro"
        Exit Sub
    
    End If
        
        
        
        

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtFatura_Change()
    If Not IsNumeric(txtFatura) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtFatura_GotFocus()
    txtFatura.SelStart = 0
    txtFatura.SelLength = 2
End Sub

Private Sub txtFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFatura)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFilial_Change()
    If Not IsNumeric(txtFilial) Then
        SendKeys "{BACKSPACE}"
    End If
    If Len(Trim$(txtFilial)) = 2 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub

Private Sub txtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilial)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtObsCanc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtObsCanc_LostFocus()
    txtObsCanc = UCase(txtObsCanc)
End Sub
