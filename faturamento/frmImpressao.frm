VERSION 5.00
Begin VB.Form frmImpressao 
   Caption         =   "Impressão de Faturas"
   ClientHeight    =   2910
   ClientLeft      =   1665
   ClientTop       =   1395
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5565
   Begin VB.Frame Frame1 
      Caption         =   "Impressão de Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtFatura 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1680
         Width           =   375
      End
      Begin VB.OptionButton optImprFatura 
         Caption         =   "Imprimir a Fatura"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optImprRelat 
         Caption         =   "Imprimir Relatório de Fatura"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Filial Fatura:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    Dim xFilialFatura As String
    
    xFilialFatura = TransFatur(txtFilial, txtFatura)
    cmdImprimir.Caption = "Aguarde ..."
    Frame1.Enabled = False
    
    If optImprRelat.Value = True Then 'se for relatório de conferencia
    
        lblCtr = "FATURA: " & xFilialFatura
        lblGravando = "Imprimindo Relatório ..."
        DoEvents
        Call imprime_fatrel(xFilialFatura)
        lblCtr = "FATURA: "
        lblGravando = "Gravando ..."
        DoEvents
        MsgBox "Registro Gravado e Enviado Relatório para Impressão. FATURA: " & xFilialFatura
    
    ElseIf optImprFatura.Value = True Then 'se for relat. de fatura
    
        If de_informa.rsSel_Fatura.State = 1 Then de_informa.rsSel_Fatura.Close
        de_informa.Sel_Fatura xFilialFatura
        
        If de_informa.rsSel_Fatura.RecordCount > 0 Then
            If de_informa.rsSel_Fatura.Fields("impresso") = "S" Then
                If MsgBox("Atenção ! Esta Fatura já Foi Impressa em " & de_informa.rsSel_Fatura.Fields("impressodata") & _
                          ". Deseja Mesmo Assim Imprimir Novamente ?", vbQuestion + vbYesNo, "Reimpressão") = vbNo Then
                    cmdImprimir.Caption = "Imprimir"
                    Frame1.Enabled = True
                    txtFilial.SetFocus
                    Exit Sub
                End If
            End If
        
            lblCtr = "FATURA: " & xFilialFatura
            lblGravando = "Imprimindo Fatura ..."
            DoEvents
            Call imprime_fat(xFilialFatura)
            de_informa.Alt_ImpressoFaturaSim xFilialFatura
            lblCtr = "FATURA: "
            lblGravando = "Gravando ..."
            DoEvents
            MsgBox "Registro Gravado e Enviado Fatura para Impressão. FATURA: " & xFilialFatura
        
        Else
            
            MsgBox "Fatura Inexistente !", vbCritical, "Erro"
        
        End If
    
    End If
    
    cmdImprimir.Caption = "Imprimir"
    Frame1.Enabled = True
    txtFilial.SetFocus

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtFatura_Change()
    If Len(Trim$(txtFatura)) > 0 Then
        If Not IsNumeric(txtFatura) Or Mid$(txtFatura, Len(txtFatura), 1) = "," Or Mid$(txtFatura, Len(txtFatura), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
End Sub

Private Sub txtFatura_GotFocus()
    txtFatura.SelStart = 0
    txtFatura.SelLength = 8
End Sub

Private Sub txtFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFatura)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFilial_Change()
    If Len(Trim$(txtFilial)) > 0 Then
        If Not IsNumeric(txtFilial) Or Mid$(txtFilial, Len(txtFilial), 1) = "," Or Mid$(txtFilial, Len(txtFilial), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
    If Len(Trim$(txtFilial)) = 2 Then
        txtFatura.SetFocus
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

Private Sub txtFilial_LostFocus()
    If Len(Trim$(txtFilial)) = 1 Then
        txtFilial = "0" & Trim$(txtFilial)
    End If
End Sub
