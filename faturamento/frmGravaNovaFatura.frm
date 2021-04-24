VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmGravaNovaFatura 
   Caption         =   "Nova Fatura - Gravação"
   ClientHeight    =   6375
   ClientLeft      =   1020
   ClientTop       =   1170
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7155
   Begin VB.Frame fraAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   720
      TabIndex        =   28
      Top             =   6120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label lblGravando 
         Alignment       =   2  'Center
         Caption         =   "Gravando ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   4335
      End
   End
   Begin VB.Frame fraAguarde2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   840
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Finalização"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame1 
         Caption         =   "PRÉ - FATURA"
         Height          =   855
         Left            =   3240
         TabIndex        =   41
         Top             =   1830
         Width           =   3735
         Begin VB.Label lblAvulsaDesc 
            Alignment       =   2  'Center
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   3465
         End
         Begin VB.Label lblPrefatura 
            Alignment       =   2  'Center
            Caption         =   "PREFATURA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   42
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   135
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   5880
         Width           =   135
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   6720
         TabIndex        =   38
         Top             =   5880
         Width           =   135
      End
      Begin VB.CommandButton cmdEditDataEmissao 
         Height          =   195
         Left            =   6720
         TabIndex        =   37
         Top             =   240
         Width           =   135
      End
      Begin VB.CheckBox chkEmissao 
         Caption         =   "Data de Emissão Editável"
         Height          =   195
         Left            =   2400
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkImprimir 
         Caption         =   "Imprimir a Fatura"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   5160
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdGravarNota 
         Caption         =   "Gravar Nova Fatura ..."
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   5040
         Width           =   2535
      End
      Begin VB.TextBox txtObsFatura 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   4
         Top             =   4560
         Width           =   6735
      End
      Begin VB.TextBox txtAcrescimo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtObsAcresc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtAbat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtObsAbat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   40
         TabIndex        =   1
         Top             =   3120
         Width           =   3615
      End
      Begin MSMask.MaskEdBox mskEmissao 
         Height          =   285
         Left            =   840
         TabIndex        =   36
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
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
         TabIndex        =   34
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "( - ) Desc. ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2220
         Width           =   1170
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
         TabIndex        =   32
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto com ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   1605
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Observações:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   4320
         Width           =   990
      End
      Begin VB.Label lblTipoAbat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3240
         TabIndex        =   26
         Top             =   2835
         Width           =   3615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   6840
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lblEmissor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5400
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Emissor:"
         Height          =   195
         Left            =   4680
         TabIndex        =   24
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3120
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblClienteCNPJ 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Top             =   600
         Width           =   1455
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
         TabIndex        =   21
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "( = )  Valor da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "( + )  Acréscimos:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "( - )  Abatimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto da Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   1545
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
         TabIndex        =   16
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   4680
         TabIndex        =   14
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblConta 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5400
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblBanco 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencto:"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmGravaNovaFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEmissao_Click()
    If chkEmissao.Value = 1 Then
        mskEmissao.Enabled = True
        mskEmissao.BackColor = xamarelo1
    Else
        mskEmissao.Enabled = False
        mskEmissao.BackColor = xbranco
    End If
End Sub

Private Sub cmdEditDataEmissao_Click()
    If chkEmissao.Visible = True Then
        chkEmissao.Visible = False
    Else
        chkEmissao.Visible = True
    End If
End Sub
Private Sub cmdGravarNota_Click()

    Dim xNumFatura As Double, xFilialFatura As String, xAvulsa As String, xUsuAbat As String, xDataAbat As Date
    Dim xUsuAcres As String, xDataAcres As Date, xFilialCtc As String
    Dim xCon As New ADODB.Connection
    Dim xrs As New ADODB.Recordset
    
    cmdGravarNota.Enabled = False
    
    fraAguarde.Top = 1800
    fraAguarde2.Top = 2040
    fraAguarde.Visible = True
    fraAguarde2.Visible = True
    
    If Not IsDate(mskEmissao) Then
        MsgBox "Data de Emissão Inválida !"
        Exit Sub
    End If
    
    If lblPrefatura = "AVULSA" Then
    Else
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura lblPrefatura
        
        If de_informa.rsSel_PreFatura.RecordCount > 0 Then
            Do Until de_informa.rsSel_PreFatura.EOF
                If de_informa.rsSel_PreFatura.Fields("data") > CDate(mskEmissao) Then
                    MsgBox "ATENÇÃO ! Não Será Possível Continuar. Existem CTCs/NFs com Data de Emissão Posterior a Data de Emissão Desta Fatura. Exemlo: " & de_informa.rsSel_PreFatura.Fields("filialctc") & ". Confira a Pré-Fatura e Verifique Se Existem Outros CTCs Com Data de Emissão Posterior a Data de Emissão Desta Fatura.", vbCritical, "ERRO"
                    Exit Sub
                End If
                de_informa.rsSel_PreFatura.MoveNext
            Loop
        End If
        
    End If
    
    If Len(Trim$(txtAbat.Text)) = 0 Then
        txtAbat.Text = 0
    End If
    If Len(Trim$(txtAcrescimo.Text)) = 0 Then
        txtAcrescimo.Text = 0
    End If
    If Len(Trim$(lblValorICMS.Caption)) = 0 Then
        lblValorICMS.Caption = 0
    End If
        
    If CDbl(SoNumeros(txtAbat)) / 100 > 0 Then
        xUsuAbat = xusuario
        xDataAbat = datahora("DATA")
    Else
        xUsuAbat = ""
        xDataAbat = "1900/01/01"
    End If
            
    If CDbl(SoNumeros(txtAcrescimo)) / 100 > 0 Then
        xUsuAcres = xusuario
        xDataAcres = datahora("DATA")
    Else
        xUsuAcres = ""
        xDataAcres = "1900/01/01"
    End If
            
    de_informa.cn_informa.BeginTrans   'inicio de transacao
    xCon.ConnectionString = xstrcon2

    xCon.ConnectionTimeout = 30
    xCon.Open
    xCon.BeginTrans
    If lblPrefatura = "AVULSA" Then
        xrs.Open "exec sp_numfat '" & frmNovaFatura.txtFilialFatura & "'", xCon, adOpenStatic, adLockBatchOptimistic
        xNumFatura = xrs.Fields(0)
        xrs.Close
        xFilialFatura = TransFatur(frmNovaFatura.txtFilialFatura, CVar(xNumFatura))
    Else
        xrs.Open "exec sp_numfat '" & Mid$(lblPrefatura, 1, 2) & "'", xCon, adOpenStatic, adLockBatchOptimistic
        xNumFatura = xrs.Fields(0)
        xrs.Close
        xFilialFatura = TransFatur(Mid$(lblPrefatura, 1, 2), CVar(xNumFatura))
    End If
    
    'busca dados do cliente (end, ie, cidade, etc)
    
    If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
    de_informa.Sel_CadCliCGC lblClienteCNPJ.Caption
    
    'grava a Fatura
    
    If (CDbl(SoNumeros(lblValorICMS.Caption)) / 100) > (CDbl(SoNumeros(lblValorFaturaBruto.Caption)) / 100) Then
        lblValorICMS.Caption = "0"
        DoEvents
    End If
    
    If lblPrefatura = "AVULSA" Then
    
        de_informa.Ins_Fatura xFilialFatura, CDate(mskEmissao.Text), datahora("HORA"), xusuario, "N", CDate(lblVencto.Caption), CDate(lblVencto.Caption), _
                              lblClienteCNPJ.Caption, lblCliente.Caption, de_informa.rsSel_CadCliCGC.Fields("ie"), de_informa.rsSel_CadCliCGC.Fields("endereco"), _
                              de_informa.rsSel_CadCliCGC.Fields("cidade"), de_informa.rsSel_CadCliCGC.Fields("uf"), de_informa.rsSel_CadCliCGC.Fields("cep"), _
                              de_informa.rsSel_CadCliCGC.Fields("endcob"), de_informa.rsSel_CadCliCGC.Fields("cidadecob"), de_informa.rsSel_CadCliCGC.Fields("ufcob"), _
                              de_informa.rsSel_CadCliCGC.Fields("cepcob"), de_informa.rsSel_CadCliCGC.Fields("telefonecob"), de_informa.rsSel_CadCliCGC.Fields("contatocob"), _
                              Mid$(lblBanco.Caption, 1, 4), Mid$(lblBanco.Caption, 6), lblConta.Caption, lblPrefatura, CDbl(SoNumeros(lblValorFaturaBrutoICMS.Caption)) / 100, _
                              CDbl(SoNumeros(lblValorICMS.Caption)) / 100, CDbl(SoNumeros(lblValorFaturaBruto.Caption)) / 100, CDbl(SoNumeros(txtAbat.Text)) / 100, _
                              lblTipoAbat.Caption, txtObsAbat.Text, xUsuAbat, xDataAbat, CDbl(SoNumeros(txtAcrescimo.Text)) / 100, txtObsAcresc.Text, xUsuAcres, _
                              xDataAcres, CDbl(SoNumeros(lblValorFatura.Caption)) / 100, txtObsFatura.Text, "", "", "N"
                              ', Trim$(lblAvulsaDesc)
    
    Else
    
        de_informa.Ins_Fatura xFilialFatura, CDate(mskEmissao.Text), datahora("HORA"), xusuario, "N", CDate(lblVencto.Caption), CDate(lblVencto.Caption), _
                              lblClienteCNPJ.Caption, lblCliente.Caption, de_informa.rsSel_CadCliCGC.Fields("ie"), de_informa.rsSel_CadCliCGC.Fields("endereco"), _
                              de_informa.rsSel_CadCliCGC.Fields("cidade"), de_informa.rsSel_CadCliCGC.Fields("uf"), de_informa.rsSel_CadCliCGC.Fields("cep"), _
                              de_informa.rsSel_CadCliCGC.Fields("endcob"), de_informa.rsSel_CadCliCGC.Fields("cidadecob"), de_informa.rsSel_CadCliCGC.Fields("ufcob"), _
                              de_informa.rsSel_CadCliCGC.Fields("cepcob"), de_informa.rsSel_CadCliCGC.Fields("telefonecob"), de_informa.rsSel_CadCliCGC.Fields("contatocob"), _
                              Mid$(lblBanco.Caption, 1, 4), Mid$(lblBanco.Caption, 6), lblConta.Caption, lblPrefatura, CDbl(SoNumeros(lblValorFaturaBrutoICMS.Caption)) / 100, _
                              CDbl(SoNumeros(lblValorICMS.Caption)) / 100, CDbl(SoNumeros(lblValorFaturaBruto.Caption)) / 100, CDbl(SoNumeros(txtAbat.Text)) / 100, _
                              lblTipoAbat.Caption, txtObsAbat.Text, xUsuAbat, xDataAbat, CDbl(SoNumeros(txtAcrescimo.Text)) / 100, txtObsAcresc.Text, xUsuAcres, _
                              xDataAcres, CDbl(SoNumeros(lblValorFatura.Caption)) / 100, txtObsFatura.Text, "", "", "N"
                              ', Trim$(lblAvulsaDesc)
        
        'grava CTCs/nfs da Fatura
    
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura lblPrefatura
           
        Do Until de_informa.rsSel_PreFatura.EOF
        
            de_informa.Ins_FaturaItem xFilialFatura, de_informa.rsSel_PreFatura.Fields("tipodoc"), de_informa.rsSel_PreFatura.Fields("filialctc"), _
                                       de_informa.rsSel_PreFatura.Fields("data"), de_informa.rsSel_PreFatura.Fields("frete"), de_informa.rsSel_PreFatura.Fields("fretebruto"), _
                                       de_informa.rsSel_PreFatura.Fields("obs"), de_informa.rsSel_PreFatura.Fields("digitacao"), de_informa.rsSel_PreFatura.Fields("emissor")
                            
            If de_informa.rsSel_PreFatura.Fields("tipodoc") = "NFS" Then
                de_informa.Alt_NFSFaturado xFilialFatura, de_informa.rsSel_PreFatura.Fields("filialctc")
            Else
                de_informa.Alt_CTCFaturado xFilialFatura, de_informa.rsSel_PreFatura.Fields("filialctc")
            End If
            
            de_informa.rsSel_PreFatura.MoveNext
        
        Loop
        
        'limpa a Prefatura
        de_informa.Excl_PreFatTudo lblPrefatura
        
    End If
    
    xCon.CommitTrans
    de_informa.cn_informa.CommitTrans        'finalizar transacao
    xCon.Close
            
    'rotina de impressão
        
    If chkImprimir.Value = 1 Then
    
        lblCtr = "FATURA: " & xFilialFatura
        lblGravando = "Imprimindo Fatura ..."
        DoEvents
        Call imprime_fat(xFilialFatura)
        de_informa.Alt_ImpressoFaturaSim xFilialFatura
        lblCtr = "FATURA: "
        lblGravando = "Gravando ..."
        DoEvents
        MsgBox "Registro Gravado e Enviado Fatura para Impressão. FATURA: " & xFilialFatura, vbInformation, "Impressão"
    
    Else
    
        lblCtr = "FATURA: "
        lblGravando = "Gravando ..."
        DoEvents
        MsgBox "Registro Gravado. FATURA: " & xFilialFatura, vbInformation, "Impressão"
    
    End If
    
    fraAguarde.Top = 3000
    fraAguarde2.Top = 3000
    fraAguarde.Visible = False
    fraAguarde2.Visible = False
    
    frmNovaFatura.txtPreFatura = ""
    frmNovaFatura.cmdGerarFat.Caption = "FATURADO"
    
    Unload Me
    
End Sub

Private Sub optImprFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optImprRelat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtAbat_Change()
    If Not IsNumeric(txtAbat) Then
        SendKeys "{BACKSPACE}"
        Exit Sub
    End If
    Call TextMoneyBox_Change(txtAbat)
    DoEvents
    If Len(Trim$(txtAbat)) > 0 Then
        If CDbl(SoNumeros(txtAbat)) / 100 > 0 Then
            txtObsAbat.Enabled = True
            txtObsAbat.BackColor = xamarelo1
            DoEvents
            Exit Sub
        End If
    End If
    txtObsAbat.Enabled = False
    txtObsAbat.BackColor = xbranco
    DoEvents

End Sub
Private Sub txtAbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAbat_LostFocus()
    If Len(Trim$(txtAbat)) > 0 Then
        If CDbl(SoNumeros(txtAbat)) / 100 > 0 Then
            If Len(Trim$(txtAcrescimo)) = 0 Then txtAcrescimo = 0
            lblValorFatura = Format((CDbl(SoNumeros(lblValorFaturaBruto)) / 100) - (CDbl(SoNumeros(txtAbat)) / 100) + (CDbl(SoNumeros(txtAcrescimo)) / 100), "##,###,##0.00")
            frmMotivosDesconto.Show 1
            Exit Sub
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
            lblValorFatura = Format((CDbl(SoNumeros(lblValorFaturaBruto)) / 100) - (CDbl(SoNumeros(txtAbat)) / 100) + (CDbl(SoNumeros(txtAcrescimo)) / 100), "##,###,##0.00")
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

Private Sub txtObsAbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtObsAbat_LostFocus()
    txtObsAbat = UCase(Trim$(txtObsAbat))
End Sub

Private Sub txtObsAcresc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtObsAcresc_LostFocus()
    txtObsAcresc = UCase(Trim$(txtObsAcresc))
End Sub

Private Sub txtObsFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtObsFatura_LostFocus()
    txtObsFatura = UCase(Trim$(txtObsFatura))
End Sub

Private Sub mskEmissao_GotFocus()
    mskEmissao.SelStart = 0
    mskEmissao.SelLength = 10
End Sub

Private Sub mskEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(mskEmissao)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub mskEmissao_LostFocus()
    If mskEmissao.Text <> "__/__/____" Then
        mskEmissao.Text = century(mskEmissao.Text)
        If IsDate(mskEmissao.Text) = False Or Mid(mskEmissao.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskEmissao.SetFocus
            Exit Sub
        End If
        If CDate(mskEmissao.Text) <> datahora("data") Then
            MsgBox "ATENÇÃO ! Confira a Data de Emissão. Data Diferente de Hoje ???", vbCritical, "Confirmação"
        End If
        If CDate(mskEmissao.Text) > datahora("data") Then
            MsgBox "ATENÇÃO ! Data de Emissão Maior que Hoje Não é Válido !!!", vbCritical, "Erro"
            mskEmissao.SetFocus
        End If
    End If
End Sub


