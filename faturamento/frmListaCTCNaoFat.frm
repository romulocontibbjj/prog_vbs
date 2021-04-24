VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaCTCNaoFat 
   Caption         =   "Relatório de CTCs Não Faturados"
   ClientHeight    =   8085
   ClientLeft      =   2490
   ClientTop       =   2460
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDados 
      Height          =   6255
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   11655
      Begin VB.CommandButton cmdGerarArq 
         Caption         =   "Gerar Arquivo ..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Imprimir Relatório ..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid gridRelNaoFat 
         Bindings        =   "frmListaCTCNaoFat.frx":0000
         Height          =   5535
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9763
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFreteLiq 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblQtde 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   5880
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Frete:"
         Height          =   195
         Left            =   7680
         TabIndex        =   25
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Qtde CTCs:"
         Height          =   195
         Left            =   4920
         TabIndex        =   24
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "Opções de Seleção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   11655
      Begin VB.Frame Frame6 
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   9585
         Begin VB.Frame Frame1 
            Caption         =   "Ordenar Por ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7080
            TabIndex        =   27
            Top             =   150
            Width           =   2415
            Begin VB.OptionButton optPorCliente 
               Caption         =   "Por Cliente+Numeração"
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Top             =   300
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton optPorNumeracao 
               Caption         =   "Por Numeração"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   540
               Width           =   2055
            End
         End
         Begin VB.OptionButton optNFS 
            Caption         =   "NFS"
            Height          =   195
            Left            =   1680
            TabIndex        =   3
            Top             =   800
            Width           =   615
         End
         Begin VB.OptionButton optCTC 
            Caption         =   "CTC"
            Height          =   195
            Left            =   720
            TabIndex        =   2
            Top             =   800
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.TextBox txtFilial 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   6360
            MaxLength       =   3
            TabIndex        =   7
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtCnpj 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   3120
            MaxLength       =   14
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdBuscaCli 
            Caption         =   "?"
            Height          =   255
            Left            =   4800
            TabIndex        =   5
            Top             =   360
            Width           =   375
         End
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
            TabIndex        =   0
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "DOC:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   800
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Filial:"
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
            Left            =   6360
            TabIndex        =   23
            Top             =   120
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            Left            =   3120
            TabIndex        =   22
            Top             =   120
            Width           =   660
         End
         Begin VB.Label lblCliente 
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   6
            Top             =   640
            Width           =   3000
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Período:"
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
            TabIndex        =   21
            Top             =   120
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1440
            TabIndex        =   19
            Top             =   360
            Width           =   90
         End
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   495
         Left            =   10080
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   10080
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmListaCTCNaoFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaCli_Click()
frm_cgc.Show
End Sub

Private Sub cmdGerarArq_Click()
    Dim xlinha As String, xFiles As String, xrs As Recordset
    
    If optCTC.Value = True Then
        If optPorCliente.Value = True Then
            Set xrs = de_informa.rsSel_CTCNaoFat
        Else
            Set xrs = de_informa.rsSel_CTCNaoFatPorNum
        End If
    ElseIf optNFS.Value = True Then
        If optPorCliente.Value = True Then
            Set xrs = de_informa.rsSel_NFSNaoFat
        Else
            Set xrs = de_informa.rsSel_NFSNaoFatPorNum
        End If
    End If

    If xrs.State <> 1 Then
        MsgBox "Dados Nao Disponíveis !"
        Exit Sub
    End If
    
    If xrs.RecordCount > 0 Then
    
        If optCTC.Value = True Then
            xFiles = "C:\INFORMA\CTCNAOFAT" & "_" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & Mid$(Trim$(CVar(Time())), 1, 2) & Mid$(Trim$(CVar(Time())), 4, 2) & ".txt"
        ElseIf optNFS.Value = True Then
            xFiles = "C:\INFORMA\NFSNAOFAT" & "_" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & Mid$(Trim$(CVar(Time())), 1, 2) & Mid$(Trim$(CVar(Time())), 4, 2) & ".txt"
        End If
        
        Open xFiles For Output As #1
        
        If optCTC.Value = True Then
        
            xlinha = "Filial-CTC#Data#Consignatario#Remetente#Destinatario#Cidade Destino#UF#Valor Merc.#Frete Liquido#Prioridade#Natureza#Observacao de Emissao#"
            
            Print #1, xlinha
        
            Do Until xrs.EOF
            
                xlinha = xrs.Fields("filialctc") & "#" & _
                         xrs.Fields("data") & "#" & _
                         xrs.Fields("respons_nome") & "#" & _
                         xrs.Fields("remet_nome") & "#" & _
                         xrs.Fields("dest_nome") & "#" & _
                         xrs.Fields("cidade_dest") & "#" & _
                         xrs.Fields("uf_dest") & "#" & _
                         xrs.Fields("valmerc") & "#" & _
                         xrs.Fields("fretefinal") & "#" & _
                         xrs.Fields("prioridade") & "#" & _
                         xrs.Fields("natureza") & "#" & _
                         xrs.Fields("obs_emissao") & "#"
                         
                Print #1, xlinha
                
                xrs.MoveNext
                
            Loop
            
        ElseIf optNFS.Value = True Then
        
            xlinha = "Filial-NFS#Data#Cliente#Correspondente#ValorLiq#"
            
            Print #1, xlinha
        
            Do Until xrs.EOF
            
                xlinha = xrs.Fields("filialnfs") & "#" & _
                         xrs.Fields("data") & "#" & _
                         xrs.Fields("cliente_nome") & "#" & _
                         xrs.Fields("corresp") & "#" & _
                         xrs.Fields("valornfsliquido") & "#"
                         
                Print #1, xlinha
                
                xrs.MoveNext
                
            Loop
        
        End If
        
        xrs.MoveFirst
        
        Close #1
        
        MsgBox "Arquivo Gerado ! " & xFiles & Chr(10) + Chr(13) + Chr(10) + Chr(13) + _
               "O Arquivo Gerado é do Tipo TEXTO Delimitado e pode ser Aberto no Excel utilizando o caracter # como delimitador.", vbInformation, "Arquivo Gerado"
        
    End If
        
End Sub

Private Sub cmdImpressao_Click()
    Dim xcont As Integer, xrs As Recordset, xsubtot As Currency, xtotger As Currency, xcontsub As Integer
    Dim xpag As Long
    
    If optNFS.Value = True Then
        MsgBox "Relatório de NFS Não Faturados em Desenvolvimento. Use a Opção Gerar Arquivo... e abra-o no Excel. (Para CTCs esta Opção já Está Pronta ...).", vbCritical, "OPS"
        Exit Sub
    End If
    
    'busca impressora para este documento
    If Dir("C:\informa.cfg") <> "" Then
        
        Open "C:\informa.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "REL" Then
                ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Exit Do
            End If
        Loop
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: REL "
        fraDados.Enabled = True
        FraPeriodo.Enabled = True
        cmdImpressao.Caption = "Imprimir Relatório ..."
        Exit Sub
        
    End If
    
    'seta impressora
    
    For Each ximpr_inst In Printers
        If ximpr_inst.DeviceName = ximpr_cfg Then
            Set Printer = ximpr_inst
            DoEvents
            Exit For
        End If
    Next
    
    xcont = 1
    xpag = 1
    xcontsub = 0
    
    If optPorCliente.Value = True Then
        Set xrs = de_informa.rsSel_CTCNaoFat
    Else
        Set xrs = de_informa.rsSel_CTCNaoFatPorNum
    End If

    If xrs.State <> 1 Then
        MsgBox "Dados Nao Disponíveis !"
        Exit Sub
    End If

    fraDados.Enabled = False
    FraPeriodo.Enabled = False
    cmdImpressao.Caption = "Aguarde Impressão"
    
    xrs.MoveFirst
    
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    
    xsubtot = 0
    xtotger = 0
        
    Do Until xrs.EOF
    
        If xcont = 1 Then
                
            Printer.FontBold = True
            Printer.Print "RELATÓRIO DE CTCs NAO FATURADOS"
            If Len(Trim$(txtCnpj)) >= 8 Then
                Printer.Print "CLIENTE: " & Mid$(lblCliente, 25)
            Else
                Printer.Print "CLIENTE: TODOS"
            End If
            Printer.Print "PERÍODO: " & mskPer1.Text & " à " & mskPer2.Text; Space(5);
            Printer.Print "Filial: " & txtFilial; Space(60);
            Printer.Print "Pág: " & zeros(xpag, 3)
            Printer.Print "----------------------------------------------------------------------------------------------------------------------"
            Printer.Print "Filial-CTC   Data     Consignatário    Remetente        Destinarário    Frete Líquido Obs_Emissão                       "
            Printer.Print "----------------------------------------------------------------------------------------------------------------------"
            Printer.FontBold = False
            xcont = 8
        End If
        
        Printer.Print xrs.Fields("filialctc"); Spc(1);
        
        Printer.Print Trim$(xrs.Fields("data")); Spc(1);
        
        Printer.Print Trim$(Mid$(xrs.Fields("respons_nome"), 1, 15)) & _
                      String(15 - Len(Trim$(Mid$(xrs.Fields("respons_nome"), 1, 15))), " "); Spc(2);
        
        Printer.Print Trim$(Mid$(xrs.Fields("remet_nome"), 1, 15)) & _
                      String(15 - Len(Trim$(Mid$(xrs.Fields("remet_nome"), 1, 15))), " "); Spc(2);
        
        Printer.Print Trim$(Mid$(xrs.Fields("dest_nome"), 1, 15)) & _
                      String(15 - Len(Trim$(Mid$(xrs.Fields("dest_nome"), 1, 15))), " "); Spc(1);
        
        Printer.Print String(13 - Len(Format(xrs.Fields("fretefinal"), "##,###,##0.00")), " "); Format(xrs.Fields("fretefinal"), "##,###,##0.00"); Spc(1);
                      
        Printer.Print Trim$(Mid$(xrs.Fields("obs_emissao"), 1, 32)) & _
                      String(32 - Len(Trim$(Mid$(xrs.Fields("obs_emissao"), 1, 32))), " ")

        xcont = xcont + 1
        xtotger = xtotger + xrs.Fields("fretefinal")
        
        xrs.MoveNext
        
        If xcont >= 90 Then
            Printer.FontBold = True
            Printer.Print "----------------------------------------------------------------------------------------------------------------------"
            xcont = 1
            xpag = xpag + 1
            Printer.FontBold = False
            Printer.NewPage
        End If
        
    Loop
    
    Printer.FontBold = True
    Printer.Print "----------------------------------------------------------------------------------------------------------------------"
    Printer.Print Space(46); "TOTAL GERAL ............"; Spc(3);
    Printer.Print String(13 - Len(Format(xtotger, "##,###,##0.00")), " "); Format(xtotger, "##,###,##0.00")
    Printer.Print "----------------------------------------------------------------------------------------------------------------------"
    Printer.FontBold = False
    
    xrs.MoveFirst

    Printer.EndDoc

    MsgBox "Dados Enviados para Impressão !", vbInformation

    fraDados.Enabled = True
    FraPeriodo.Enabled = True
    cmdImpressao.Caption = "Imprimir Relatório ..."
    mskPer1.SetFocus

End Sub

Private Sub cmdProcessar_Click()
    
    Dim xfreteliq As Currency, xrs As Recordset
    
    If Trim$(txtCnpj) = "" Then
        txtCnpj.Text = "%"
    Else
        If InStr(Len(txtCnpj.Text), txtCnpj.Text, "%", vbTextCompare) Then
            txtCnpj.Text = txtCnpj.Text
        Else
            txtCnpj.Text = txtCnpj.Text & "%"
        End If
    End If
    
    If Trim$(txtFilial) <> "%" Then
        txtFilial = txtFilial & "%"
    End If
    
    If optPorCliente.Value = False And optPorNumeracao = False Then
        optPorCliente.Value = True
    End If
    
    If optCTC.Value = False And optNFS = False Then
        optCTC.Value = True
    End If
    
    cmdProcessar.Caption = "Aguarde ..."
    
    fraDados.Enabled = False
    FraPeriodo.Enabled = False
    cmdProcessar.Enabled = False
    Me.MousePointer = 11
    DoEvents
    Set gridRelNaoFat.DataSource = de_informa
    gridRelNaoFat.ClearFields
    
    If optCTC.Value = True Then
    
        If optPorCliente.Value = True Then
            
            If de_informa.rsSel_CTCNaoFat.State = 1 Then de_informa.rsSel_CTCNaoFat.Close
            de_informa.Sel_CTCNaoFat CDate(mskPer1), CDate(mskPer2), txtCnpj, txtFilial
            
            gridRelNaoFat.DataMember = "Sel_CTCNaoFat"
            gridRelNaoFat.Refresh
            Set xrs = de_informa.rsSel_CTCNaoFat
            
        Else
        
            If de_informa.rsSel_CTCNaoFatPorNum.State = 1 Then de_informa.rsSel_CTCNaoFatPorNum.Close
            de_informa.Sel_CTCNaoFatPorNum CDate(mskPer1), CDate(mskPer2), txtCnpj, txtFilial
            
            gridRelNaoFat.DataMember = "Sel_CTCNaoFatPorNum"
            gridRelNaoFat.Refresh
            Set xrs = de_informa.rsSel_CTCNaoFatPorNum
        
        End If
        
    ElseIf optNFS.Value = True Then
    
        If optPorCliente.Value = True Then
            
            If de_informa.rsSel_NFSNaoFat.State = 1 Then de_informa.rsSel_NFSNaoFat.Close
            de_informa.Sel_NFSNaoFat CDate(mskPer1), CDate(mskPer2), txtCnpj, txtFilial
            
            gridRelNaoFat.DataMember = "Sel_NFSNaoFat"
            gridRelNaoFat.Refresh
            Set xrs = de_informa.rsSel_NFSNaoFat
            
        Else
        
            If de_informa.rsSel_NFSNaoFatPorNum.State = 1 Then de_informa.rsSel_NFSNaoFatPorNum.Close
            de_informa.Sel_NFSNaoFatPorNum CDate(mskPer1), CDate(mskPer2), txtCnpj, txtFilial
            
            gridRelNaoFat.DataMember = "Sel_NFSNaoFatPorNum"
            gridRelNaoFat.Refresh
            Set xrs = de_informa.rsSel_NFSNaoFatPorNum
        
        End If
    
    End If
        
        xfreteliq = 0
        
        Do Until xrs.EOF
            If optCTC.Value = True Then
                If IsNull(xrs.Fields("fretefinal")) = True Then
                    xfreteliq = xfreteliq + 0
                Else
                    xfreteliq = xfreteliq + xrs.Fields("fretefinal")
                End If
            ElseIf optNFS.Value = True Then
                xfreteliq = xfreteliq + xrs.Fields("valornfsliquido")
            End If
            xrs.MoveNext
        Loop
        
        If xrs.RecordCount < 1 Then
            MsgBox "Não Há Dados a Serem Exibidos !", vbInformation
            txtCnpj.Text = Left(txtCnpj.Text, Len(txtCnpj.Text) - 1)
            cmdGerarArq.Enabled = False
            cmdImpressao.Enabled = False
        Else
            cmdGerarArq.Enabled = True
            cmdImpressao.Enabled = True
            xrs.MoveFirst
        End If
        
        lblQtde = xrs.RecordCount
        lblFreteLiq = Format(xfreteliq, "##,###,##0.00")
        
    cmdProcessar.Caption = "Processar"
    fraDados.Enabled = True
    FraPeriodo.Enabled = True
    cmdProcessar.Enabled = True
    Me.MousePointer = 0
    DoEvents
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mdiFatura.ToolFaturamento.Visible = False
    If de_informa.rsSel_CTCNaoFat.State = 1 Then de_informa.rsSel_CTCNaoFat.Close
    If de_informa.rsSel_CTCNaoFatPorNum.State = 1 Then de_informa.rsSel_CTCNaoFatPorNum.Close
    gridRelNaoFat.DataMember = "Sel_CTCNaoFat"
    gridRelNaoFat.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFatura.ToolFaturamento.Visible = True
End Sub

Private Sub optCTC_Click()
    If de_informa.rsSel_CTCNaoFat.State = 1 Then de_informa.rsSel_CTCNaoFat.Close
    If de_informa.rsSel_CTCNaoFatPorNum.State = 1 Then de_informa.rsSel_CTCNaoFatPorNum.Close
    If de_informa.rsSel_NFSNaoFat.State = 1 Then de_informa.rsSel_NFSNaoFat.Close
    If de_informa.rsSel_NFSNaoFatPorNum.State = 1 Then de_informa.rsSel_NFSNaoFatPorNum.Close
    gridRelNaoFat.DataMember = ""
    gridRelNaoFat.Refresh
    lblQtde = ""
    lblFreteLiq = ""
End Sub

Private Sub optNFS_Click()
    If de_informa.rsSel_CTCNaoFat.State = 1 Then de_informa.rsSel_CTCNaoFat.Close
    If de_informa.rsSel_CTCNaoFatPorNum.State = 1 Then de_informa.rsSel_CTCNaoFatPorNum.Close
    If de_informa.rsSel_NFSNaoFat.State = 1 Then de_informa.rsSel_NFSNaoFat.Close
    If de_informa.rsSel_NFSNaoFatPorNum.State = 1 Then de_informa.rsSel_NFSNaoFatPorNum.Close
    gridRelNaoFat.DataMember = ""
    gridRelNaoFat.Refresh
    lblQtde = ""
    lblFreteLiq = ""

End Sub
Private Sub optPorCliente_Click()
    If de_informa.rsSel_CTCNaoFat.State = 1 Then de_informa.rsSel_CTCNaoFat.Close
    If de_informa.rsSel_CTCNaoFatPorNum.State = 1 Then de_informa.rsSel_CTCNaoFatPorNum.Close
    If de_informa.rsSel_NFSNaoFat.State = 1 Then de_informa.rsSel_NFSNaoFat.Close
    If de_informa.rsSel_NFSNaoFatPorNum.State = 1 Then de_informa.rsSel_NFSNaoFatPorNum.Close
    gridRelNaoFat.DataMember = ""
    gridRelNaoFat.Refresh
    lblQtde = ""
    lblFreteLiq = ""

End Sub
Private Sub optPorNumeracao_Click()
    If de_informa.rsSel_CTCNaoFat.State = 1 Then de_informa.rsSel_CTCNaoFat.Close
    If de_informa.rsSel_CTCNaoFatPorNum.State = 1 Then de_informa.rsSel_CTCNaoFatPorNum.Close
    If de_informa.rsSel_NFSNaoFat.State = 1 Then de_informa.rsSel_NFSNaoFat.Close
    If de_informa.rsSel_NFSNaoFatPorNum.State = 1 Then de_informa.rsSel_NFSNaoFatPorNum.Close
    gridRelNaoFat.DataMember = ""
    gridRelNaoFat.Refresh
    lblQtde = ""
    lblFreteLiq = ""

End Sub

Private Sub txtFilial_Change()
    If Len(Trim$(txtFilial)) > 0 Then
        If Not IsNumeric(txtFilial) Or Mid$(txtFilial, Len(txtFilial), 1) = "," Or Mid$(txtFilial, Len(txtFilial), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
    If Len(Trim$(txtFilial)) = 2 Then
        cmdProcessar.SetFocus
    End If
End Sub

Private Sub txtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub
Private Sub mskPer1_GotFocus()
    mskPer1.SelStart = 0
    mskPer1.SelLength = 10
End Sub
Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskPer1_LostFocus()
    If mskPer1.Text <> "__/__/____" Then
        mskPer1.Text = century(mskPer1.Text)
        If IsDate(mskPer1.Text) = False Or Mid(mskPer1.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub mskPer2_GotFocus()
    mskPer2.SelStart = 0
    mskPer2.SelLength = 10
End Sub
Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskPer2_LostFocus()
    If mskPer2.Text <> "__/__/____" Then
        mskPer2.Text = century(mskPer2.Text)
        If IsDate(mskPer2.Text) = False Or Mid(mskPer2.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
    End If
End Sub


Private Sub txtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub
