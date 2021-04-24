VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPreFatPend 
   Caption         =   "Pré-Faturas Pendentes"
   ClientHeight    =   6090
   ClientLeft      =   510
   ClientTop       =   1455
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9765
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdAlteraVencPrefat 
      Caption         =   "Altera Vencto Pré-Fatura"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
   End
   Begin VB.OptionButton optTodasPre 
      Caption         =   "Todos as Pré-Faturas"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton optUsuarioPre 
      Caption         =   "Somente o Usuario:"
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VB.CommandButton cmdExcluirPreFat 
      Caption         =   "Excluir Pré-Fatura Pendente ..."
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdGerarFatura 
      Caption         =   "Gerar Fatura Final ..."
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdConsPreFat 
      Caption         =   "Consultar / Alterar Pré-Fatura ..."
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid gridPreFat 
      Bindings        =   "frmPreFatPend.frx":0000
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
      _Version        =   393216
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
      DataMember      =   "Sel_PreFatPend"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "filialprefatura"
         Caption         =   "Pré-Fatura"
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
         DataField       =   "emissor"
         Caption         =   "Emissor"
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
         DataField       =   "vencimento"
         Caption         =   "Vencimento"
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
         DataField       =   "cliente_nome"
         Caption         =   "Cliente"
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
         DataField       =   "totvalor"
         Caption         =   "Valor Pré-Fatura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3780,284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1484,787
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPreFatPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualiza_Click()

End Sub

Private Sub cmdAlteraVencPrefat_Click()
    
    frmAlteraVencPrefat.lblPrefat = gridPreFat.Columns(0)
    frmAlteraVencPrefat.lblCliente = gridPreFat.Columns(3)
    frmAlteraVencPrefat.lblVencAtual = gridPreFat.Columns(2)
    frmAlteraVencPrefat.lblValor = gridPreFat.Columns(4)
    frmAlteraVencPrefat.lblUsu = gridPreFat.Columns(1)
    frmAlteraVencPrefat.Show 1
    
    If optTodasPre.Value = True Then
        If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
        de_informa.Sel_PreFatPend "%"
        gridPreFat.DataMember = "sel_prefatpend"
        gridPreFat.Refresh
    ElseIf optUsuarioPre.Value = True Then
        If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
        de_informa.Sel_PreFatPend xusuario
        gridPreFat.DataMember = "sel_prefatpend"
        gridPreFat.Refresh
    End If
    
    

End Sub

Private Sub cmdConsPreFat_Click()
    
    Me.Hide
    frmNovaFatura.txtPreFilial.Enabled = True
    frmNovaFatura.txtPreFilial.BackColor = xamarelo1
    frmNovaFatura.txtPreFatura.Enabled = True
    frmNovaFatura.txtPreFatura.BackColor = xamarelo1
    frmNovaFatura.cmdBuscaPreFat.Enabled = True
    frmNovaFatura.tabTipoFatura.TabEnabled(0) = True
    frmNovaFatura.tabTipoFatura.TabEnabled(1) = False
    frmNovaFatura.tabTipoFatura.TabEnabled(2) = False
    frmNovaFatura.txtPreFilial = Mid$(gridPreFat.Columns(0), 1, 2)
    frmNovaFatura.txtPreFatura = Mid$(gridPreFat.Columns(0), 3, 6)
    frmNovaFatura.Show
    Unload Me

End Sub
Private Sub cmdExcluirPreFat_Click()
    
    If MsgBox("Confirma a Exclusão da Pré-Fatura " & gridPreFat.Columns(0) & " ?", vbQuestion + vbYesNo, "Confirma Exclusão") = vbYes Then
    
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura gridPreFat.Columns(0)
    
        de_informa.cn_informa.BeginTrans
    
            If de_informa.rsSel_PreFatura.Fields("tipodoc") = "NFS" Then
                de_informa.Alt_LimpaFaturaNFS gridPreFat.Columns(0)
            Else
                de_informa.Alt_LimpaFaturaCTC gridPreFat.Columns(0)
            End If
        
            de_informa.Excl_PreFatTudo gridPreFat.Columns(0)
        
        de_informa.cn_informa.CommitTrans
        
        If optTodasPre.Value = True Then
            If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
            de_informa.Sel_PreFatPend "%"
            gridPreFat.DataMember = "sel_prefatpend"
            gridPreFat.Refresh
        ElseIf optUsuarioPre.Value = True Then
            If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
            de_informa.Sel_PreFatPend xusuario
            gridPreFat.DataMember = "sel_prefatpend"
            gridPreFat.Refresh
        End If
            
    End If

End Sub

Private Sub cmdGerarFatura_Click()
    If MsgBox("Você Confirma a Geração de FATURA para a Pré-Fatura Número " & gridPreFat.Columns(0) & " ?", vbYesNo + vbQuestion, "Confirma") = vbYes Then
            
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura gridPreFat.Columns(0)
        
        If de_informa.rsSel_PreFatura.RecordCount < 1 Then
            MsgBox "Esta Pré-Fatura não Consta mais no Banco de Dados. Provavelmente Alguém já a Faturou ou Excluiu !"
        Else
            
            frmGravaNovaFatura.lblPrefatura = gridPreFat.Columns(0)
            frmGravaNovaFatura.lblClienteCNPJ.Caption = de_informa.rsSel_PreFatura.Fields("cliente_cgc")
            frmGravaNovaFatura.lblCliente.Caption = de_informa.rsSel_PreFatura.Fields("cliente_nome")
            frmGravaNovaFatura.lblVencto.Caption = de_informa.rsSel_PreFatura.Fields("vencimento")
            frmGravaNovaFatura.lblEmissor.Caption = de_informa.rsSel_PreFatura.Fields("emissor")
            frmGravaNovaFatura.lblBanco.Caption = zeros(de_informa.rsSel_PreFatura.Fields("banco"), 4) & "-" & de_informa.rsSel_PreFatura.Fields("banconome")
            frmGravaNovaFatura.lblConta.Caption = de_informa.rsSel_PreFatura.Fields("conta")
            
            If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
            de_informa.Sel_PreFaturaTotais gridPreFat.Columns(0)
            
            frmGravaNovaFatura.lblValorFaturaBrutoICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
            frmGravaNovaFatura.lblValorFaturaBruto.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            frmGravaNovaFatura.lblValorICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot") - de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            frmGravaNovaFatura.lblValorFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            
            frmGravaNovaFatura.mskEmissao.Mask = ""
            frmGravaNovaFatura.mskEmissao.Text = datahora("DATA")
            frmGravaNovaFatura.mskEmissao.Mask = "##/##/####"
            
            frmGravaNovaFatura.Show 1
            
        End If
        
        If optTodasPre.Value = True Then
            If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
            de_informa.Sel_PreFatPend "%"
            gridPreFat.DataMember = "sel_prefatpend"
            gridPreFat.Refresh
        ElseIf optUsuarioPre.Value = True Then
            If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
            de_informa.Sel_PreFatPend xusuario
            gridPreFat.DataMember = "sel_prefatpend"
            gridPreFat.Refresh
        End If
    
    End If
            
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    
    'fazendo por pré-fatura
    If de_informa.rsSel_AcertoNFS.State = 1 Then de_informa.rsSel_AcertoNFS.Close
    de_informa.Sel_AcertoNFS
    
    Do Until de_informa.rsSel_AcertoNFS.EOF
        de_informa.Alt_AcertoNFS de_informa.rsSel_AcertoNFS.Fields("filialprefatura"), de_informa.rsSel_AcertoNFS.Fields("filialctc")
        de_informa.rsSel_AcertoNFS.MoveNext
    Loop
    
    MsgBox "fim"
    
End Sub

Private Sub Form_Load()

    optUsuarioPre.Caption = "Somente o Usuario: " & xusuario
    
    ''65
    'contador = contador + 6
    'If Mid$(xdireitos, contador, 1) = "0" Then
    '    mnuPrePend.Enabled = False
    '    ToolFaturamento.Buttons(2).Enabled = False
    'End If
    
    '66
    contador = 66
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmPreFatPend.optTodasPre.Enabled = False
    End If
    
    '67
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmPreFatPend.optUsuarioPre.Enabled = False
    End If
    
    '68
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmPreFatPend.cmdAlteraVencPrefat.Enabled = False
    End If
                    
    '69
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmPreFatPend.cmdConsPreFat.Enabled = False
    End If
    '70
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmPreFatPend.cmdExcluirPreFat.Enabled = False
    End If
    
    '71
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmPreFatPend.cmdGerarFatura.Enabled = False
    End If
    
    '72
    'contador = contador + 1
    'If Mid$(xdireitos, contador, 1) = "0" Then
    '    mnuConsultaFat.Enabled = False
    '   ToolFaturamento.Buttons(3).Enabled = False
    'End If
    
    
    If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
    de_informa.Sel_PreFatPend xusuario
    
    gridPreFat.DataMember = "sel_prefatpend"
    gridPreFat.Refresh
    
    
    
    
    
End Sub

Private Sub optTodasPre_Click()
    If optTodasPre.Value = True Then
        If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
        de_informa.Sel_PreFatPend "%"
        gridPreFat.DataMember = "sel_prefatpend"
        gridPreFat.Refresh
    ElseIf optUsuarioPre.Value = True Then
        If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
        de_informa.Sel_PreFatPend xusuario
        gridPreFat.DataMember = "sel_prefatpend"
        gridPreFat.Refresh
    End If
End Sub

Private Sub optUsuarioPre_Click()
    If optTodasPre.Value = True Then
        If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
        de_informa.Sel_PreFatPend "%"
        gridPreFat.DataMember = "sel_prefatpend"
        gridPreFat.Refresh
    ElseIf optUsuarioPre.Value = True Then
        If de_informa.rsSel_PreFatPend.State = 1 Then de_informa.rsSel_PreFatPend.Close
        de_informa.Sel_PreFatPend xusuario
        gridPreFat.DataMember = "sel_prefatpend"
        gridPreFat.Refresh
    End If
End Sub
