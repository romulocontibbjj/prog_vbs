VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiInforma 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Informa - Módulo de Informação - Intec Cargo - V2.5"
   ClientHeight    =   8250
   ClientLeft      =   735
   ClientTop       =   1590
   ClientWidth     =   12375
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiInforma.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Timer tmAlarmeUrgencia 
      Left            =   240
      Top             =   2160
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1535
      ButtonWidth     =   2037
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta SAC"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ocorr / POD"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acompanh."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ctrl. Canhotos"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alarme"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Estatística 1"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta MNF"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "An. Ocorr."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "An. Entregas"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "usuário ativo no momento"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13758
            MinWidth        =   13758
            Object.ToolTipText     =   "Observações"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2/12/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":6936
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":6E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":6F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":70A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":722A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":7BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":7D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":82CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":86DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":8886
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":8C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":94DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInforma.frx":9A12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivos 
      Caption         =   "Arquivos"
      Begin VB.Menu mnuImpEDI 
         Caption         =   "Importação PROCEDA/EDI"
      End
      Begin VB.Menu mnuExportEDI 
         Caption         =   "Exportação PROCEDA/EDI"
      End
      Begin VB.Menu mnuExpSipla 
         Caption         =   "Exportação SITLA"
      End
      Begin VB.Menu mnusepara1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Log de Usuários"
      End
      Begin VB.Menu mnuDupl 
         Caption         =   "CTC/CTR Duplicados"
      End
   End
   Begin VB.Menu mnuCad 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuOcorr 
         Caption         =   "Ocorrências"
      End
      Begin VB.Menu mnuCadUsu 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuCadCli 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuFeriados 
         Caption         =   "Feriados"
      End
      Begin VB.Menu mnuPrazos 
         Caption         =   "Prazos de Entrega"
      End
   End
   Begin VB.Menu mnuProcesso 
      Caption         =   "Processos"
      Begin VB.Menu mnuSac 
         Caption         =   "Consulta SAC"
      End
      Begin VB.Menu mnuConsMnf 
         Caption         =   "Consulta Manifestos"
      End
      Begin VB.Menu mnuPOD 
         Caption         =   "Baixas e Ocorrências Manuais (POD)"
      End
      Begin VB.Menu mnuAcompanha 
         Caption         =   "Acompanhamento de Clientes/Região"
      End
      Begin VB.Menu mnuAcompInf 
         Caption         =   "Acompanhamento da Informação (Resumo)"
      End
      Begin VB.Menu mnuGeraProtCanhoto 
         Caption         =   "Controle de Canhotos de Notas Fiscais"
      End
      Begin VB.Menu mnuDevolucoes 
         Caption         =   "Controle de Devoluções"
      End
      Begin VB.Menu mnuproclin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlarme 
         Caption         =   "Alarme Urgencias e Prioridades"
      End
      Begin VB.Menu mnuproclin2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDtEmissaoNF 
         Caption         =   "Processa Arquivos Exclusivos para o Cliente"
      End
      Begin VB.Menu mnuAverba 
         Caption         =   "Processa Arquivo da Pamcary (Averbação)"
      End
      Begin VB.Menu mnuenvemail 
         Caption         =   "Informar Ocorrências para Cliente (Email)"
      End
      Begin VB.Menu mnuCanc 
         Caption         =   "Cancelar CTC"
      End
      Begin VB.Menu mnuproclin3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecalPrazo 
         Caption         =   "Recalcular os Prazos de Entrega ..."
      End
      Begin VB.Menu mnuPrevEntrega 
         Caption         =   "Recalcular as Datas de Previsão de Entrega ..."
      End
      Begin VB.Menu lin66 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsSacVL 
         Caption         =   "Consulta SAC Modelo Videolar"
      End
      Begin VB.Menu mnuVideolar 
         Caption         =   "Controle Específico VIDEOLAR"
      End
      Begin VB.Menu lin90 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCotarFrete 
         Caption         =   "Cotação de Frete"
      End
   End
   Begin VB.Menu mnuInformacao 
      Caption         =   "Gerencial"
      Begin VB.Menu mnuAn_Entr 
         Caption         =   "Análise de Entregas"
      End
      Begin VB.Menu mnuAn_Ocorr 
         Caption         =   "Análise de Ocorrências"
      End
      Begin VB.Menu mnuAn_Oper 
         Caption         =   "Análise Estatística"
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuprot 
         Caption         =   "Protocolo para Arquivo..."
         Begin VB.Menu mnuimprprot 
            Caption         =   "Impressão do Protocolo p/ Arquivo (CTCs)"
         End
         Begin VB.Menu mnuProtocolo 
            Caption         =   "Reimpressão de Protocolo Já Impresso"
         End
      End
      Begin VB.Menu mnuRelatArq 
         Caption         =   "Gerar Relatórios / Arquivos"
      End
   End
   Begin VB.Menu mnuColeta 
      Caption         =   "Coletas"
      Begin VB.Menu mnuColetaAcomp 
         Caption         =   "Acompanhamento de Coletas"
      End
      Begin VB.Menu mnuCancColeta 
         Caption         =   "Cancelar Coleta"
      End
      Begin VB.Menu mnuConsultaColeta 
         Caption         =   "Consulta Coleta"
      End
      Begin VB.Menu mnuColetaOrdem 
         Caption         =   "Ordem de Coleta"
      End
      Begin VB.Menu mnuPODColeta 
         Caption         =   "POD Coleta"
      End
      Begin VB.Menu Lin67 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColetaIMP 
         Caption         =   "Definir Impressora"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "mdiInforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Activate()
    Dim dia As String
    StatusBar1.Panels.Item(1).Text = "USR: " & xusuario
    StatusBar1.Panels.Item(3).Text = diasemana(datahora("data"))
    xamarelo1 = &HC0FFFF
    xamarelo2 = &HFFFF&
    xbranco = &H8000000E

    
    
End Sub

Private Sub MDIForm_Load()
    xusuario = ""
'    If App.PrevInstance = True Then
'        MsgBox "ATENÇÃO ! O Informa já está aberto. Não é possível ter dois Sistemas Informa em execução simultaneamente.", vbCritical, "ERRO"
'        Unload frmAcesso
'        Unload mdiInforma
'        End
'    Else
        frmAcesso.Show 1
'    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If xusuario <> "" Then
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "LOGOFF", xusuario, "OK"
    End If

    End
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set mdiInforma = Nothing
End Sub

Private Sub mnuAcompanha_Click()
    If Mid$(xdireitos, 15, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmAcompanha.Show
    End If
End Sub

Private Sub mnuAcompInf_Click()
    If Mid$(xdireitos, 27, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmAcompInformacao.Show
    End If

    
End Sub

Private Sub mnuAlarme_Click()
    Dim xtm_Interval As Long
    xtm_Interval = mdiInforma.tmAlarmeUrgencia.Interval
    mdiInforma.tmAlarmeUrgencia.Interval = 0
    frmAlarmeUrg.Show 1
    mdiInforma.tmAlarmeUrgencia.Interval = xtm_Interval
End Sub

Private Sub mnuAn_Entr_Click()
    If Mid$(xdireitos, 18, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmEscCliPer.Caption = "Análise de Entregas"
        frmEscCliPer.optAir.Enabled = True
        frmEscCliPer.optRodo.Enabled = True
        frmEscCliPer.fraAnalise.Visible = True
        'frmAnEntregas.Show
        frmEscCliPer.Show
    End If
End Sub

Private Sub mnuAn_Ocorr_Click()
    If Mid$(xdireitos, 19, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmEscCliPer.Caption = "Análise de Ocorrências"
        frmEscCliPer.chkModal.Enabled = True
        'frmEscCliPer.optAir.Enabled = True
        'frmEscCliPer.optRodo.Enabled = True
        frmEscCliPer.chkModal = 1
        frmEscCliPer.Show
    End If
End Sub

Private Sub mnuAn_Oper_Click()
    If Mid$(xdireitos, 20, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmEscCliPer.Caption = "Análise Estatística"
        frmEscCliPer.chkModal.Enabled = True
        frmEscCliPer.chkModal = 1
        frmEscCliPer.TxtFilial.BackColor = xamarelo1
        frmEscCliPer.TxtFilial.Enabled = True
        frmEscCliPer.lblFilial.Enabled = True
        frmEscCliPer.Show
        'frmAnEstat.Show
    End If
End Sub
Private Sub mnuCadCli_Click()
    frmCadCli.Show
End Sub
Private Sub mnuCadUsu_Click()
    If Mid$(xdireitos, 4, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmCadUsu.Show 1
    End If
End Sub
Private Sub mnuCanc_Click()
    If Mid$(xdireitos, 16, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmCancCTC.Show 1
    End If
End Sub

Private Sub mnuConfig_Click()
'    frmAcertoAWS.Show
'    If Mid$(xdireitos, 24, 1) = "0" Then
'        MsgBox "Acesso Não Permitido !"
'    Else
'        frmConfig.Show 1
'    End If
End Sub

Private Sub mnuCancColeta_Click()
    If Mid$(xdireitos, 35, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
    frmCancelaColeta.Show 1
    End If
End Sub

Private Sub mnuColetaAcomp_Click()
    If Mid$(xdireitos, 34, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
    frmListaColeta.Show
    End If
End Sub

Private Sub mnuColetaIMP_Click()
frmColetaImpressoras.Show 1
End Sub

Private Sub mnuColetaOrdem_Click()
    If Mid$(xdireitos, 32, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
    frmOrdemColeta.Show
    End If
End Sub

Private Sub mnuConsMnf_Click()
    If Mid$(xdireitos, 10, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmMNF.Show
    End If
End Sub

Private Sub mnuConsSacVL_Click()
If Mid$(xdireitos, 40, 1) = "0" Then
    MsgBox "Acesso Não Permitido !"
Else
    frmVLSac.Show 1
    End If
End Sub

Private Sub mnuConsultaColeta_Click()
frmConsultaColeta.Show 1
End Sub

Private Sub mnuCotarFrete_Click()
If Mid$(xdireitos, 46, 1) = "0" Then
    MsgBox "Acesso Não Permitido !"
Else
    'frmVideoLarCtr.Show 1
End If
End Sub

Private Sub mnuDevolucoes_Click()
    If Mid$(xdireitos, 29, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
'        frmDevolução.Show 1
    End If
End Sub

Private Sub mnuDtEmissaoNF_Click()
    If Mid$(xdireitos, 17, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmDtEmissaoNF.Show 1
    End If
End Sub

Private Sub mnuEmail_Click()
    frmEnvEmail.Show 1
End Sub

Private Sub mnuenvemail_Click()
    If Mid$(xdireitos, 14, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        'frmEnvEmail.Show 1
    End If
End Sub

Private Sub mnuExportEDI_Click()
    If Mid$(xdireitos, 25, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmExportEDI.Show 1
    End If
End Sub

Private Sub mnuExpSipla_Click()
    If Mid$(xdireitos, 2, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmExportSitla.Show 1
    End If
End Sub
Private Sub mnuFeriados_Click()
    frmCadFeriados.Show
End Sub

Private Sub mnuGeraProtCanhoto_Click()
    If Mid$(xdireitos, 28, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmGeraProtCanhotos.Show 1
    End If
End Sub

Private Sub mnuImpEDI_Click()
    If Mid$(xdireitos, 1, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmEdiImport.Show
    End If
End Sub

Private Sub mnuimprprot_Click()
    If Mid$(xdireitos, 21, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        If MsgBox("Deseja Imprimir o Relatório de CTCs Físicos Baixados, para Envio dos Documentos para o Arquivo ? (PROTOCOLO)", vbQuestion + vbYesNo, "Confirmação de Relatório") = vbYes Then
            mdiInforma.StatusBar1.Panels.Item(2).Text = "AGUARDE IMPRESSAO DO RELATORIO ..."
            DoEvents
            Call rel_arquivo
            mdiInforma.StatusBar1.Panels.Item(2).Text = ""
            DoEvents
        End If
    End If
End Sub

Private Sub mnuLog_Click()
    If Mid$(xdireitos, 26, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmLogUsuarios.Show 1
    End If

End Sub

Private Sub mnuOcorr_Click()
    frmCadOcorr.Show
End Sub
Private Sub mnuPod_Click()
    If Mid$(xdireitos, 11, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmPod.Show
    End If
End Sub

Private Sub mnuPODColeta_Click()
    If Mid$(xdireitos, 33, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
    frmPodColeta.Show 1
    End If
End Sub

Private Sub mnuPrazos_Click()
    If Mid$(xdireitos, 8, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmCadPrazos.Show
    End If
End Sub
Private Sub mnuPrevEntrega_Click()
    If Mid$(xdireitos, 13, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmRecalcPrevEntr.Show
    End If
End Sub
Private Sub mnuProtocolo_Click()
    If Mid$(xdireitos, 21, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmReimprProt.Show 1
    End If
End Sub
Private Sub mnuRecalPrazo_Click()
    If Mid$(xdireitos, 12, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmRecalcPrazos.Show
    End If
End Sub

Private Sub mnuRelatArq_Click()
    frmRelatEspecificos.Show
End Sub

Private Sub mnuSac_Click()
    If Mid$(xdireitos, 10, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmSac.Show
    End If
End Sub
Private Sub mnuSair_Click()
    
    Unload Me
    
End Sub

Private Sub mnuVideolar_Click()
If Mid$(xdireitos, 43, 1) = "0" Then
    MsgBox "Acesso Não Permitido !"
Else
    frmVideoLarCtr.Show 1
End If
End Sub

Private Sub tmAlarmeUrgencia_Timer()
    xtempoalarme = xtempoalarme + 1
    If xtempoalarme >= 60 Or mdiInforma.tmAlarmeUrgencia.Interval = 15000 Then
        mdiInforma.tmAlarmeUrgencia.Interval = 60000
        If de_informa.rsSel_Urgencias.State = 1 Then de_informa.rsSel_Urgencias.Close
        de_informa.Sel_Urgencias
        If de_informa.rsSel_Prioridades.State = 1 Then de_informa.rsSel_Prioridades.Close
        de_informa.Sel_Prioridades
        If (de_informa.rsSel_Urgencias.RecordCount + de_informa.rsSel_Prioridades.RecordCount) > 0 _
        And Int(mdiInforma.tmAlarmeUrgencia.Interval) > 0 Then
            de_informa.ins_LogUsuario "ALARME", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
            mdiInforma.tmAlarmeUrgencia.Interval = 0
            frmAlarmeUrg.Show 1
            mdiInforma.tmAlarmeUrgencia.Interval = 60000
        End If
        xtempoalarme = 0
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        mnuSac_Click
    ElseIf Button.Index = 2 Then
        mnuPod_Click
    ElseIf Button.Index = 3 Then
        mnuAcompanha_Click
    ElseIf Button.Index = 4 Then
        mnuGeraProtCanhoto_Click
    ElseIf Button.Index = 5 Then
        mnuAlarme_Click
    ElseIf Button.Index = 6 Then
        mnuAn_Oper_Click
    ElseIf Button.Index = 7 Then
        mnuConsMnf_Click
    ElseIf Button.Index = 8 Then
        mnuAn_Ocorr_Click
    ElseIf Button.Index = 9 Then
        mnuAn_Entr_Click
    ElseIf Button.Index = 10 Then
        mnuSair_Click
    End If
End Sub
