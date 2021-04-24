VERSION 5.00
Begin VB.Form frmSelGerarArqAcomp 
   Caption         =   "Gerar Arquivo - Acompanhamento de Cliente"
   ClientHeight    =   5640
   ClientLeft      =   3570
   ClientTop       =   1410
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   4455
   Begin VB.Frame fraSelecao 
      Caption         =   "Seleção dos Dados (SEM POSIÇÃO)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdSair 
         Caption         =   "S A I R"
         Height          =   495
         Left            =   2640
         TabIndex        =   18
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar Arquivo..."
         Height          =   495
         Left            =   600
         TabIndex        =   11
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Frame fraSubContr 
         Caption         =   "4 - Transportador Sub-Contratado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   3975
         Begin VB.ComboBox cmbSubContr 
            Height          =   315
            Left            =   1320
            TabIndex        =   10
            Text            =   "Todos"
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Transportador:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1035
         End
      End
      Begin VB.Frame fraUf 
         Caption         =   "2 - Região"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
         Begin VB.Frame Frame1 
            Caption         =   "Localidade"
            Height          =   1095
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   3735
            Begin VB.TextBox txtCidade 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   17
               Top             =   720
               Width           =   2175
            End
            Begin VB.OptionButton optCidade 
               Caption         =   "Cidade que Comece com:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1440
               TabIndex        =   16
               Top             =   360
               Width           =   2175
            End
            Begin VB.OptionButton optInterior 
               Caption         =   "Interior"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton optCapital 
               Caption         =   "Capital"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.CheckBox chkTodoEstado 
            Caption         =   "Todo o Estado"
            Height          =   255
            Left            =   2280
            TabIndex        =   12
            Top             =   360
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.ComboBox cmbUf 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Text            =   "Todos"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Estado/UF:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame fraModal 
         Caption         =   "1 - Modal"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3975
         Begin VB.TextBox txtFilial 
            Height          =   285
            Left            =   3480
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox chkModal 
            Caption         =   "Todos Modais"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.OptionButton optAir 
            Caption         =   "Aéreo"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2160
            TabIndex        =   3
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optRodo 
            Caption         =   "Rodoviário"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2160
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmSelGerarArqAcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkModal_Click()
    If chkModal.Value = 0 Then
        optRodo.Enabled = True
        optAir.Enabled = True
    Else
        optRodo.Enabled = False
        optAir.Enabled = False
    End If
End Sub

Private Sub Option3_Click()

End Sub

Private Sub chkTodoEstado_Click()
    If chkTodoEstado.Value = 1 Then
        optCapital.Enabled = False
        optInterior.Enabled = False
        optCidade.Enabled = False
        txtCidade.Enabled = False
        optCapital.Value = True
        txtCidade.Text = ""
    Else
        optCapital.Enabled = True
        optInterior.Enabled = True
        optCidade.Enabled = True
    End If
End Sub

Private Sub cmdGerar_Click()
    Dim xdata1 As Date, xdata2 As Date, xcgc As String, xcgcdest As String, xmodal As String, xuf As String
    Dim xregiao As String, xCidade As String, xtranspsub As String, xregiaosac As String
    Dim xremet_nome As String, xremet_cgc As String, xnumnf As String, xfilialctc As String
    Dim xdata_ctc As Date, xcidade_dest As String, xuf_dest As String, xmodaltransp As String
    Dim xocorr As String, xrecebedor As String, xDigitado As Date, xprioridade As String
    
    Me.MousePointer = 11
    
    If frmAcompanha.optPorEmissao.Value = True Then  'por emissao
        If frmAcompanha.optPer15d.Value = True Then
            xdata1 = datahora("data") - 15
            xdata2 = datahora("data")
        ElseIf frmAcompanha.opt30d.Value = True Then
            xdata1 = datahora("data") - 30
            xdata2 = datahora("data")
        ElseIf frmAcompanha.opt60d.Value = True Then
            xdata1 = datahora("data") - 60
            xdata2 = datahora("data")
        Else
            MsgBox "Período Escolhido Inválido !"
            Exit Sub
        End If
    End If
    If frmAcompanha.optPorMes.Value = True Then   'por mes
        xdata1 = CDate(Mid$(frmAcompanha.comboMesAnoAcomp.ItemData(frmAcompanha.comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(frmAcompanha.comboMesAnoAcomp.ItemData(frmAcompanha.comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 "01")
        xdata2 = CDate(Mid$(frmAcompanha.comboMesAnoAcomp.ItemData(frmAcompanha.comboMesAnoAcomp.ListIndex), 1, 4) & "/" & _
                 Mid$(frmAcompanha.comboMesAnoAcomp.ItemData(frmAcompanha.comboMesAnoAcomp.ListIndex), 5, 2) & "/" & _
                 UltDiaMes(Val(Mid$(frmAcompanha.comboMesAnoAcomp.ItemData(frmAcompanha.comboMesAnoAcomp.ListIndex), 5, 2)), _
                           Val(Mid$(frmAcompanha.comboMesAnoAcomp.ItemData(frmAcompanha.comboMesAnoAcomp.ListIndex), 1, 4))))
                           
        If xdata2 > datahora("DATA") Then xdata2 = datahora("DATA")
        
    End If
    If frmAcompanha.optPorPeriodo.Value = True Then   'por periodo
        xdata1 = CDate(frmAcompanha.mskPer1)
        xdata2 = CDate(frmAcompanha.mskPer2)
    End If
    
    If frmAcompanha.optSelReg = True Then
        xregiaosac = frmAcompanha.txtregiaosac
        xcgc = "%"
        xcgcdest = "%"
    Else
        If frmAcompanha.optRemetente = True Then
            xcgc = frmAcompanha.TxtCGCRem & "%"
            If xcgc = "%%" Then xcgc = "%"
            xcgcdest = "%"
        Else
            xcgc = "%"
            xcgcdest = frmAcompanha.TxtCGCRem & "%"
            If xcgcdest = "%%" Then xcgcdest = "%"
        End If
        xregiaosac = "%"
    End If
    
    'MODAL
    If chkModal.Value = 1 Then
        xmodal = "%"
    Else
        If optRodo.Value = True Then
            xmodal = "RODOVIARIO%"
        Else
            xmodal = "AEREO%"
        End If
    End If
    
    'ESTADO
    If cmbUf.Text = "Todos" Then
        xuf = "%"
    Else
        xuf = Trim$(cmbUf.Text) & "%"
    End If
    
    'REGIAO OU CIDADE
    If chkTodoEstado.Value = 1 Then
        xregiao = "%"
        xCidade = "%"
    Else
        If optCapital.Value = True Then
            xregiao = "CAPITAL%"
            xCidade = "%"
        ElseIf optInterior.Value = True Then
            xregiao = "INTERIOR%"
            xCidade = "%"
        ElseIf optCidade.Value = True Then
            xregiao = "%"
            xCidade = txtCidade.Text & "%"
        End If
    End If
    
    'TRANSP. SUBCONTRATADA
    If cmbSubContr.Text = "Todos" Then
        xtranspsub = "%"
    ElseIf cmbSubContr.Text = "? (INTEC/REPRES)" Then
        xtranspsub = "?%"
    Else
        xtranspsub = Trim$(cmbSubContr.Text) & "%"
    End If
    
    
     'PRIORIDADES
'    If frmAcompanha.optPrioriTodo = True Then
'        xprioridade = "%"
'    ElseIf frmAcompanha.optPrioriPrioridades = True Then
'        xprioridade = "PRIORIDADE%"
'    ElseIf frmAcompanha.optPrioriUrgencias = True Then
'        xprioridade = "URGÊNCIA%"
'    End If
    

    'BUSCA OS DADOS
    
    'PARA GERAÇÃO DOS DADOS SEM POSIÇÃO
    
    'POR NF
    If fraSelecao.Caption = "Seleção dos Dados (SEM POSIÇÃO) - POR NF" Then
    
        If de_informa.rsSel_GeraArqSemPos.State = 1 Then de_informa.rsSel_GeraArqSemPos.Close
        de_informa.Sel_GeraArqSemPos xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
    
        If de_informa.rsSel_GeraArqSemPos.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\SEMPOSICAO.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Obs. de Emissao;"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqSemPos.EOF
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqSemPos.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqSemPos.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqSemPos.Fields("remet_cgc"), 13, 2) & ";" & _
                        de_informa.rsSel_GeraArqSemPos.Fields("numnf") & ";" & Mid$(de_informa.rsSel_GeraArqSemPos.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqSemPos.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqSemPos.Fields("data") & ";" & de_informa.rsSel_GeraArqSemPos.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqSemPos.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqSemPos.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqSemPos.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqSemPos.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqSemPos.Fields("transp_sub") & ";" & _
                        de_informa.rsSel_GeraArqSemPos.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqSemPos.Fields("obs_emissao") & ";"
                Print #1, xlinha
                de_informa.rsSel_GeraArqSemPos.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-SEM POSIÇÃO"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo SEMPOSICAO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqSemPos.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
        
    'POR CTC
    ElseIf fraSelecao.Caption = "Seleção dos Dados (SEM POSIÇÃO) - POR CTC" Then
    
        If de_informa.rsSel_GeraArqSemPosCTC.State = 1 Then de_informa.rsSel_GeraArqSemPosCTC.Close
        de_informa.Sel_GeraArqSemPosCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
    
        If de_informa.rsSel_GeraArqSemPosCTC.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\SEMPOSICAO.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqSemPosCTC.EOF
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqSemPosCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqSemPosCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqSemPosCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                        Mid$(de_informa.rsSel_GeraArqSemPosCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqSemPosCTC.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqSemPosCTC.Fields("data") & ";" & de_informa.rsSel_GeraArqSemPosCTC.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqSemPosCTC.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqSemPosCTC.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqSemPosCTC.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqSemPosCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqSemPosCTC.Fields("transp_sub") & ";" & _
                        de_informa.rsSel_GeraArqSemPosCTC.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqSemPosCTC.Fields("nfs") & ";" & _
                        de_informa.rsSel_GeraArqSemPosCTC.Fields("obs_emissao") & ";"
                Print #1, xlinha
                de_informa.rsSel_GeraArqSemPosCTC.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-SEM POSIÇÃO"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo SEMPOSICAO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqSemPosCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
        
    'GERACAO DE DADOS EM OCORRÊNCIA
    
    'POR NF
    ElseIf fraSelecao.Caption = "Seleção dos Dados (EM OCORRÊNCIA) - POR NF" Then
    
        If de_informa.rsSel_GeraArqOcorr.State = 1 Then de_informa.rsSel_GeraArqOcorr.Close
        de_informa.Sel_GeraArqOcorr xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
    
        If de_informa.rsSel_GeraArqOcorr.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\EM_OCORRENCIA.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Obs. de Emissao;Ocorrências"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqOcorr.EOF
                'busca ocorrências deste ctc
                xocorr = ""
                If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                de_informa.Sel_ConsOcorr2 de_informa.rsSel_GeraArqOcorr.Fields("filialctc"), "01"
                Do Until de_informa.rsSel_ConsOcorr2.EOF
                    xdataocorr = de_informa.rsSel_ConsOcorr2.Fields("data")
                    xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
                    xocorr = xocorr & "(" & xdataocorr & "-" & Trim$(de_informa.rsSel_ConsOcorr2.Fields("descr_ocorr")) & ") - "
                    de_informa.rsSel_ConsOcorr2.MoveNext
                Loop
                xocorr = Mid$(xocorr, 1, Len(xocorr) - 3)
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqOcorr.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqOcorr.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqOcorr.Fields("remet_cgc"), 13, 2) & ";" & _
                        de_informa.rsSel_GeraArqOcorr.Fields("numnf") & ";" & Mid$(de_informa.rsSel_GeraArqOcorr.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqOcorr.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqOcorr.Fields("data") & ";" & de_informa.rsSel_GeraArqOcorr.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqOcorr.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqOcorr.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqOcorr.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqOcorr.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqOcorr.Fields("transp_sub") & ";" & _
                        de_informa.rsSel_GeraArqOcorr.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqOcorr.Fields("obs_emissao") & ";" & xocorr
                Print #1, xlinha
                de_informa.rsSel_GeraArqOcorr.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM OCORRÊNCIA"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo EM_OCORRENCIA.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqOcorr.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
    
    'POR CTC
    ElseIf fraSelecao.Caption = "Seleção dos Dados (EM OCORRÊNCIA) - POR CTC" Then
    
        If de_informa.rsSel_GeraArqOcorrCTC.State = 1 Then de_informa.rsSel_GeraArqOcorrCTC.Close
        de_informa.Sel_GeraArqOcorrCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
    
        If de_informa.rsSel_GeraArqOcorrCTC.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\EM_OCORRENCIA.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;Ocorrências"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqOcorrCTC.EOF
                'busca ocorrências deste ctc
                xocorr = ""
                If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                de_informa.Sel_ConsOcorr2 de_informa.rsSel_GeraArqOcorrCTC.Fields("filialctc"), "01"
                Do Until de_informa.rsSel_ConsOcorr2.EOF
                    xdataocorr = de_informa.rsSel_ConsOcorr2.Fields("data")
                    xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
                    xocorr = xocorr & "(" & xdataocorr & "-" & Trim$(de_informa.rsSel_ConsOcorr2.Fields("descr_ocorr")) & ") - "
                    de_informa.rsSel_ConsOcorr2.MoveNext
                Loop
                If Len(xocorr) > 3 Then
                    xocorr = Mid$(xocorr, 1, Len(xocorr) - 3)
                End If
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqOcorrCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqOcorrCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqOcorrCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                        Mid$(de_informa.rsSel_GeraArqOcorrCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqOcorrCTC.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqOcorrCTC.Fields("data") & ";" & de_informa.rsSel_GeraArqOcorrCTC.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqOcorrCTC.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqOcorrCTC.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqOcorrCTC.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqOcorrCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqOcorrCTC.Fields("transp_sub") & ";" & _
                        de_informa.rsSel_GeraArqOcorrCTC.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqOcorrCTC.Fields("nfs") & ";" & _
                        de_informa.rsSel_GeraArqOcorrCTC.Fields("obs_emissao") & ";" & xocorr
                Print #1, xlinha
                de_informa.rsSel_GeraArqOcorrCTC.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM OCORRÊNCIA"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo EM_OCORRENCIA.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqOcorrCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
    
    'GERAÇÃO DE DADOS EM TRANSITO
    
    'POR NF
    ElseIf fraSelecao.Caption = "Seleção dos Dados (EM TRÂNSITO) - POR NF" Then
    
        If de_informa.rsSel_GeraArqTransito.State = 1 Then de_informa.rsSel_GeraArqTransito.Close
        de_informa.Sel_GeraArqTransito xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
    
        If de_informa.rsSel_GeraArqTransito.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\EM_TRANSITO.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Obs. de Emissao;"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqTransito.EOF
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqTransito.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqTransito.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqTransito.Fields("remet_cgc"), 13, 2) & ";" & _
                        de_informa.rsSel_GeraArqTransito.Fields("numnf") & ";" & Mid$(de_informa.rsSel_GeraArqTransito.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqTransito.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqTransito.Fields("data") & ";" & de_informa.rsSel_GeraArqTransito.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqTransito.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqTransito.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqTransito.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqTransito.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqTransito.Fields("transp_sub") & ";" & _
                        de_informa.rsSel_GeraArqTransito.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqTransito.Fields("obs_emissao") & ";"
                Print #1, xlinha
                de_informa.rsSel_GeraArqTransito.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM TRÂNSITO"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo EM_TRANSITO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqTransito.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
    
    'POR CTC
    ElseIf fraSelecao.Caption = "Seleção dos Dados (EM TRÂNSITO) - POR CTC" Then
    
        If de_informa.rsSel_GeraArqTransitoCTC.State = 1 Then de_informa.rsSel_GeraArqTransitoCTC.Close
        de_informa.Sel_GeraArqTransitoCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
    
        If de_informa.rsSel_GeraArqTransitoCTC.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\EM_TRANSITO.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqTransitoCTC.EOF
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqTransitoCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqTransitoCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqTransitoCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                        Mid$(de_informa.rsSel_GeraArqTransitoCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqTransitoCTC.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqTransitoCTC.Fields("data") & ";" & de_informa.rsSel_GeraArqTransitoCTC.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqTransitoCTC.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqTransitoCTC.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqTransitoCTC.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqTransitoCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqTransitoCTC.Fields("transp_sub") & ";" & _
                        de_informa.rsSel_GeraArqTransitoCTC.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqTransitoCTC.Fields("nfs") & ";" & _
                        de_informa.rsSel_GeraArqTransitoCTC.Fields("obs_emissao") & ";"
                Print #1, xlinha
                de_informa.rsSel_GeraArqTransitoCTC.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-EM TRÂNSITO"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo EM_TRANSITO.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqTransitoCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
    
    
    'GERAÇÃO DE DADOS ENTREGUES
    
    'POR NF
    ElseIf fraSelecao.Caption = "Seleção dos Dados (ENTREGUE) - POR NF" Then
        
        If de_informa.rsSel_GeraArqEntregue.State = 1 Then de_informa.rsSel_GeraArqEntregue.Close

        de_informa.Sel_GeraArqEntregue xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial

        If de_informa.rsSel_GeraArqEntregue.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\ENTREGUE.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Num.NF;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Data Entrega;Hora Entrega;Recebedor;Transp.Sub;Destinatario;Obs. de Emissao;Entrega Digitada Em;"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqEntregue.EOF
                If IsNull(de_informa.rsSel_GeraArqEntregue.Fields("receb")) Then
                    xrecebedor = de_informa.rsSel_GeraArqEntregue.Fields("recebpre")
                Else
                    xrecebedor = de_informa.rsSel_GeraArqEntregue.Fields("receb")
                End If
                xlinha = ""
                xDigitado = CDate(Year(de_informa.rsSel_GeraArqEntregue.Fields("usu_datapre")) & "/" & _
                            Month(de_informa.rsSel_GeraArqEntregue.Fields("usu_datapre")) & "/" & _
                            Day(de_informa.rsSel_GeraArqEntregue.Fields("usu_datapre")))
                xlinha = de_informa.rsSel_GeraArqEntregue.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqEntregue.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqEntregue.Fields("remet_cgc"), 13, 2) & ";" & _
                        de_informa.rsSel_GeraArqEntregue.Fields("numnf") & ";" & Mid$(de_informa.rsSel_GeraArqEntregue.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqEntregue.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqEntregue.Fields("data") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqEntregue.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqEntregue.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("dtentrega") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("hsentrega") & ";" & _
                        xrecebedor & ";" & de_informa.rsSel_GeraArqEntregue.Fields("transp_sub") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqEntregue.Fields("obs_emissao") & ";" _
                        & xDigitado & ";"
                Print #1, xlinha
                de_informa.rsSel_GeraArqEntregue.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-ENTREGUE"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo ENTREGUE.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqEntregue.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
    
    'POR CTC
    ElseIf fraSelecao.Caption = "Seleção dos Dados (ENTREGUE) - POR CTC" Then
        
        If de_informa.rsSel_GeraArqEntregueCTC.State = 1 Then de_informa.rsSel_GeraArqEntregueCTC.Close

        de_informa.Sel_GeraArqEntregueCTC xdata1, xdata2, xcgc, xcgcdest, xregiaosac, xmodal, xuf, xregiao, xCidade, xtranspsub, xprioridade, txtFilial
        
        If de_informa.rsSel_GeraArqEntregueCTC.RecordCount < 1 Then
            Me.MousePointer = 0
            MsgBox "Não Há Dados há serem gerados a partir das opções selecionadas !"
        Else
            Open "C:\ENTREGUE.TXT" For Output As #1
            'cria cabeçário do arquivo (campos)
            xlinha = "Cliente Remet.;CGC Cliente;Filial-CTC;Data CTC;Prioridade;Cidade Dest.;UF;Modal;Prev.Entrega;Data Entrega;Hora Entrega;Recebedor;Transp.Sub;Destinatario;Notas Fiscais;Obs. de Emissao;Entrega Digitada Em"
            Print #1, xlinha
            Do Until de_informa.rsSel_GeraArqEntregueCTC.EOF
                If IsNull(de_informa.rsSel_GeraArqEntregueCTC.Fields("receb")) Then
                    xrecebedor = de_informa.rsSel_GeraArqEntregueCTC.Fields("recebpre")
                Else
                    xrecebedor = de_informa.rsSel_GeraArqEntregueCTC.Fields("receb")
                End If
                xDigitado = CDate(Year(de_informa.rsSel_GeraArqEntregueCTC.Fields("usu_datapre")) & "/" & _
                            Month(de_informa.rsSel_GeraArqEntregueCTC.Fields("usu_datapre")) & "/" & _
                            Day(de_informa.rsSel_GeraArqEntregueCTC.Fields("usu_datapre")))
                xlinha = ""
                xlinha = de_informa.rsSel_GeraArqEntregueCTC.Fields("remet_nome") & ";" & Mid$(de_informa.rsSel_GeraArqEntregueCTC.Fields("remet_cgc"), 1, 12) & "-" & Mid$(de_informa.rsSel_GeraArqEntregueCTC.Fields("remet_cgc"), 13, 2) & ";" & _
                        Mid$(de_informa.rsSel_GeraArqEntregueCTC.Fields("filialctc"), 1, 2) & "-" & Mid$(de_informa.rsSel_GeraArqEntregueCTC.Fields("filialctc"), 3, 8) & ";" & _
                        de_informa.rsSel_GeraArqEntregueCTC.Fields("data") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("prioridade") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("cidade_dest") & ";" & _
                        de_informa.rsSel_GeraArqEntregueCTC.Fields("uf_dest") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("modal") & ";" & _
                        de_informa.rsSel_GeraArqEntregueCTC.Fields("prev_entrega") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("dtentrega") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("hsentrega") & ";" & _
                        xrecebedor & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("transp_sub") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("dest_nome") & ";" & de_informa.rsSel_GeraArqEntregueCTC.Fields("nfs") & ";" & _
                        de_informa.rsSel_GeraArqEntregueCTC.Fields("obs_emissao") & ";" & xDigitado & ";"
                Print #1, xlinha
                de_informa.rsSel_GeraArqEntregueCTC.MoveNext
            Loop
            Close #1
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "PROCESSO", xusuario, "ACOMP. DE CLIENTE: GERAÇÃO DE ARQUIVO-ENTREGUE"
            
            Me.MousePointer = 0
            MsgBox "OK ! Dados Gerados." & Chr(13) & Chr(10) & _
                    "Arquivo ENTREGUE.TXT gravado no diretório  C:\ de seu micro." & Chr(13) & Chr(10) & _
                    "Quantidade de Registros com as opções selecionadas: " & de_informa.rsSel_GeraArqEntregueCTC.RecordCount & " NFs." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Para abrir este arquivo no EXCEL: " & Chr(13) & Chr(10) & _
                    "Abra o EXCEL, Escolha ARQUIVO/ABRIR. Na Janela ABRIR e na opção ARQUIVOS DO TIPO escolha ARQUIVOS DE TEXTO e certifique-se de estar no diretório C:\ . Selecione o arquivo SEMPOSICAO.TXT e pressione o botão ABRIR desta janela." & Chr(13) & Chr(10) & _
                    "Na Janela ASSISTENTE DE IMPORTAÇÃO DE TEXTO escolha a opção DELIMITADO e clique em AVANÇAR. Como delimitadores, selecione somente a opção PONTO E VÍRGULA e não deixe mais nenhuma opção selecionada. Clique em CONCLUIR desta janela.", vbOKOnly
        End If
    
    End If
    
    Me.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Set frmSelGerarArqAcomp = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    'preenche combo de UF
    If de_informa.rsSel_Ufs.State = 1 Then de_informa.rsSel_Ufs.Close
    de_informa.Sel_Ufs
    de_informa.rsSel_Ufs.MoveFirst
    cmbUf.AddItem "Todos"
    Do Until de_informa.rsSel_Ufs.EOF
        cmbUf.AddItem de_informa.rsSel_Ufs.Fields("uf")
        de_informa.rsSel_Ufs.MoveNext
    Loop
    'preenche combo de SubContratados
    If de_informa.rsSel_SubContratados.State = 1 Then de_informa.rsSel_SubContratados.Close
    de_informa.Sel_SubContratados
    de_informa.rsSel_SubContratados.MoveFirst
    cmbSubContr.AddItem "Todos"
    cmbSubContr.AddItem "? (INTEC/REPRES)"
    Do Until de_informa.rsSel_SubContratados.EOF
        cmbSubContr.AddItem de_informa.rsSel_SubContratados.Fields("transportador")
        de_informa.rsSel_SubContratados.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSelGerarArqAcomp = Nothing
End Sub

Private Sub optCidade_Click()
    If optCidade.Value = True Then
        txtCidade.Enabled = True
        txtCidade.SetFocus
    Else
        txtCidade.Enabled = False
    End If
End Sub
