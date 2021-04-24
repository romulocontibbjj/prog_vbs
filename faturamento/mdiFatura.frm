VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiFatura 
   BackColor       =   &H00400000&
   Caption         =   "Sistema Informa - Módulo de Faturamento - V.1.92"
   ClientHeight    =   6270
   ClientLeft      =   2460
   ClientTop       =   2025
   ClientWidth     =   9795
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6015
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7410
            MinWidth        =   7410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8821
            MinWidth        =   8821
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "28/06/2005"
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
      Left            =   360
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ToolFaturamento 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   1111
      ButtonWidth     =   2381
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nova Pre-Fatura"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pre-Fatura Pend."
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta Fatura"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Arq. Pre-Fatura"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Arq. Fatura"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivos 
      Caption         =   "Arquivos"
      Begin VB.Menu mnuFatEletronica 
         Caption         =   "Fatura Eletrônica"
      End
      Begin VB.Menu mnuPreFatura 
         Caption         =   "Arquivo de Pré-Fatura"
      End
      Begin VB.Menu mnuConfBona 
         Caption         =   "Arquivo de Conferência BONAGURA"
      End
      Begin VB.Menu mnuArqFaturasBONAGURA 
         Caption         =   "Arquivo de Faturas BONAGURA"
      End
      Begin VB.Menu mnulin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfImpressoras 
         Caption         =   "Configuração de Impressoras"
      End
   End
   Begin VB.Menu mnuFaturamento 
      Caption         =   "Faturamento / Cobrança"
      Begin VB.Menu mnuNovaFat 
         Caption         =   "Nova Pré-Fatura / Fatura Avulsa"
      End
      Begin VB.Menu mnuPreFatCons 
         Caption         =   "Pré-Fatura Consulta / Altera"
      End
      Begin VB.Menu mnuPrePend 
         Caption         =   "Pré-Faturas Pendentes / Gerar Fatura"
      End
      Begin VB.Menu mnulin2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaFat 
         Caption         =   "Consulta Faturas (Prorrog/Quitacao/Abat)"
      End
      Begin VB.Menu mnuCancelamento 
         Caption         =   "Cancelar Faturas"
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Impressão de Faturas"
      End
   End
   Begin VB.Menu mnuRelat 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelFatura 
         Caption         =   "Faturamento"
      End
      Begin VB.Menu mnuEmAberto 
         Caption         =   "Faturas em Aberto"
      End
      Begin VB.Menu mnuRelCobra 
         Caption         =   "Cobrança"
      End
      Begin VB.Menu mnuNaoFaturado 
         Caption         =   "Movimento Não Faturado"
      End
      Begin VB.Menu mnuEtiq 
         Caption         =   "Etiquetas Endereço"
      End
      Begin VB.Menu mnulin29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGerencial 
         Caption         =   "Informações Gerenciais"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "mdiFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
    StatusBar1.Panels.Item(1).Text = "CNX:" & xCnx & "  BCO:" & xBco & "  USR:" & xusuario
    StatusBar1.Panels.Item(3).Text = diasemana(datahora("data"))
End Sub

Private Sub MDIForm_Load()
        
    xusuario = ""
    xamarelo1 = &HC0FFFF
    xamarelo2 = &HFFFF&
    xamarelo3 = &H80FFFF
    xbranco = &H8000000E
    frmAcesso.Show 1
'    xstrcon = "driver={SQL Server};server=192.9.205.3;database=intec;user id=sa;password=zx11bbb7"
'    xstrcon = "driver={SQL Server};server=cassio;database=intec"


'   Validar Acessos lendo a Variárel xdireitos(Global)-Franklin
    
        '49
        contador = 49
        If Mid$(xdireitos, contador, 1) = "0" Then
            'tirarAcesso = True
            mnuArquivos.Enabled = False
            ToolFaturamento.Buttons(4).Enabled = False
            ToolFaturamento.Buttons(5).Enabled = False
        
        End If
        
        '50
        contador = contador + 1
        If Mid$(xdireitos, contador, 1) = "0" Then
            'tirarAcesso = True
            mnuFaturamento.Enabled = False
            ToolFaturamento.Buttons(1).Enabled = False
            ToolFaturamento.Buttons(2).Enabled = False
            ToolFaturamento.Buttons(3).Enabled = False
        
        End If
                '51
                contador = contador + 1
                If Mid$(xdireitos, contador, 1) = "0" Then
                    mnuNovaFat.Enabled = False
                    ToolFaturamento.Buttons(1).Enabled = False
                End If
                
                'De 52 até 58 será validado no Load do Form frmNovaFatura
                    
                '59
                contador = contador + 8
                If Mid$(xdireitos, contador, 1) = "0" Then
                    mnuPreFatCons.Enabled = False
                End If
                    
                  'De 60 até 64 será validado no Load do Form frmNovaFatura
                    
                '65
                contador = contador + 6
                If Mid$(xdireitos, contador, 1) = "0" Then
                    mnuPrePend.Enabled = False
                    ToolFaturamento.Buttons(2).Enabled = False
                End If
                    
                    '66
                    contador = contador + 1
                    If Mid$(xdireitos, contador, 1) = "0" Then
                        frmPreFatPend.optTodasPre.Enabled = False
                    End If
                    
                    '67
                    contador = 67
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
                    contador = contador + 1
                    If Mid$(xdireitos, contador, 1) = "0" Then
                        mnuConsultaFat.Enabled = False
                        ToolFaturamento.Buttons(3).Enabled = False
                    End If
                    
                    ' De 73 até 76 esta no form
                    
                    '77
                    contador = contador + 5
                    If Mid$(xdireitos, contador, 1) = "0" Then
                        mnuCancelamento.Enabled = False
                    End If
                    
                        '78
                        contador = contador + 1
                        If Mid$(xdireitos, contador, 1) = "0" Then
                            frmCancelamento.cmdCancelar.Enabled = False
                        End If
                    
                    '79
                    contador = contador + 1
                    If Mid$(xdireitos, contador, 1) = "0" Then
                        mnuReimprime.Enabled = False
                    End If
                    
                        '80
                        contador = contador + 1
                        If Mid$(xdireitos, contador, 1) = "0" Then
                            frmImpressao.cmdImprimir.Enabled = False
                        End If
                                  
                    '81
                    contador = contador + 1
                    If Mid$(xdireitos, contador, 1) = "0" Then
                        mnuRelat.Enabled = False
                    End If
                
                        '82
                        contador = contador + 1
                        If Mid$(xdireitos, contador, 1) = "0" Then
                            mnuRelFatura.Enabled = False
                        End If
                    
                        '83
                        contador = contador + 1
                        If Mid$(xdireitos, contador, 1) = "0" Then
                            mnuEmAberto.Enabled = False
                        End If
                    
                        '84
                        contador = contador + 1
                        If Mid$(xdireitos, contador, 1) = "0" Then
                            mnuNaoFaturado.Enabled = False
                        End If
                    
                        '85
                        contador = contador + 1
                        If Mid$(xdireitos, contador, 1) = "0" Then
                            mnuEtiq.Enabled = False
                        End If
'-Franklin
                        
End Sub
Function soma(X As Integer)

    Dim somatoria As Integer
        
    somatoria = X + 1
    
    soma = somatoria

End Function
Private Sub mnuCancPre_Click()

End Sub

Private Sub mnuArqFaturasBONAGURA_Click()
FRM_GerarArquivoFatura.Show 1
End Sub

Private Sub mnuCancelamento_Click()
frmCancelamento.Show 1
End Sub
Private Sub mnuConfBona_Click()
    frmConfBonagura.Show 1
End Sub

Private Sub mnuConfImpressoras_Click()
    frmControleImpressoras.Show 1
End Sub
Private Sub mnuConsultaFat_Click()
    frmConsultaFatura.Show
End Sub
Private Sub mnuEmAberto_Click()
    frmEmAberto.Show 1
End Sub
Private Sub mnuEtiq_Click()
    frmImprEtiquetas.Show 1
End Sub
Private Sub mnuExpEDI_Click()
    frmExportaEDI.Show 1
End Sub
Private Sub mnuFatEletronica_Click()
    frmFaturaEletronica.Show 1
End Sub

Private Sub mnuGerencial_Click()
    frmGerencial.Show
End Sub

Private Sub mnuNaoFaturado_Click()
    frmListaCTCNaoFat.Show
End Sub
Private Sub mnuNovaFat_Click()
    frmNovaFatura.Show
End Sub
Private Sub mnuPreFatCons_Click()
    frmNovaFatura.txtPreFilial.Enabled = True
    frmNovaFatura.txtPreFilial.BackColor = xamarelo1
    frmNovaFatura.txtPreFatura.Enabled = True
    frmNovaFatura.txtPreFatura.BackColor = xamarelo1
    frmNovaFatura.cmdBuscaPreFat.Enabled = True
    frmNovaFatura.tabTipoFatura.TabEnabled(0) = True
    frmNovaFatura.tabTipoFatura.TabEnabled(1) = False
    frmNovaFatura.tabTipoFatura.TabEnabled(2) = False
    frmNovaFatura.Show
End Sub
Private Sub mnuPreFatura_Click()
    frmArqPreFatura.Show 1
End Sub
Private Sub mnuPrePend_Click()
    frmPreFatPend.Show 1
End Sub
Private Sub mnuReimprime_Click()
    frmImpressao.Show 1
End Sub
Private Sub mnuRelFatura_Click()
    frmListaFaturas.Show
End Sub
Private Sub mnuSair_Click()
    End
    Unload Me
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
End Sub
Private Sub ToolFaturamento_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        mnuNovaFat_Click
    ElseIf Button.Index = 2 Then
        mnuPrePend_Click
    ElseIf Button.Index = 3 Then
        mnuConsultaFat_Click
    ElseIf Button.Index = 4 Then
        mnuPreFatura_Click
    ElseIf Button.Index = 5 Then
        mnuFatEletronica_Click
    ElseIf Button.Index = 6 Then
        mnuSair_Click
    End If
End Sub
