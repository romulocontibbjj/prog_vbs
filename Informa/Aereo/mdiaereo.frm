VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiAereo 
   BackColor       =   &H00000000&
   Caption         =   "Emissão - Módulo Aéreo"
   ClientHeight    =   7950
   ClientLeft      =   360
   ClientTop       =   2835
   ClientWidth     =   12450
   Icon            =   "mdiaereo.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivos"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuCiaAerea 
         Caption         =   "Cia. Aérea"
      End
      Begin VB.Menu mnuFormulario 
         Caption         =   "Formulários AWB Cia Aérea"
      End
      Begin VB.Menu mnuTabPrecos 
         Caption         =   "Tabela Preços Cia Aérea"
         Begin VB.Menu mnu_CadTabPrecoINCLUSAO 
            Caption         =   "Cadastrar Novas Tabelas"
         End
         Begin VB.Menu mnu_CadTabPrecoALTERACAO_UNIT 
            Caption         =   "Reajustar Tabelas Cadastradas"
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "Localidades/Destinos"
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes Remet/Dest"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRepres 
         Caption         =   "Representantes Intec"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadProdIATA 
         Caption         =   "Categoria IATA de Produtos"
      End
      Begin VB.Menu mnuProdutosINT 
         Caption         =   "Categoria Interna de Produtos "
      End
      Begin VB.Menu mnuEspecie 
         Caption         =   "Espécie Embalagem"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObs 
         Caption         =   "Observações Padrão"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Consultas 
      Caption         =   "Processos"
      Begin VB.Menu mnu_consultaAWB 
         Caption         =   "Consultar AWB"
      End
      Begin VB.Menu mnu_acomp 
         Caption         =   "Acompanhamento de AWBs"
      End
      Begin VB.Menu mnu_voo 
         Caption         =   "Inserir / Alterar Vôos de AWBs"
      End
   End
   Begin VB.Menu mnuEmissoes 
      Caption         =   "Emissões"
      Begin VB.Menu mnuAwb 
         Caption         =   "Conhecimento Aéreo (AWB)"
      End
      Begin VB.Menu mnuManifesto 
         Caption         =   "Manifestos"
      End
      Begin VB.Menu mnuLote 
         Caption         =   "Etiquetas de Lote"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVolume 
         Caption         =   "Etiquetas de Volume"
      End
   End
   Begin VB.Menu mnuRelat 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuReportCia 
         Caption         =   "Relatório para Cia. Aérea"
      End
      Begin VB.Menu mnuAirAcompanhaREL 
         Caption         =   "Acompanhamento de Emissões"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Configurações"
      Begin VB.Menu mnuPrinters 
         Caption         =   "Impressoras"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "mdiAereo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Activate()
    If Dir("c:\printer.cfg") = "" Then
    MsgBox "Você não possui o arquivo de configuração de impressoras. Antes de continuar, é imprescindível que você configure as configure.", vbExclamation, "IMPRESSORAS"
    frmControleImpressoras.Show 1
    End If
End Sub

Private Sub MDIForm_Load()
xAmarelo = &HC0FFFF
xBranco = &H80000014
xAzul = &H800000
xPreto = &H0&
xLaranja = &HC0E0FF
xCinzaClaro = &HE0E0E0
Leave = False


'If Mid(StringDireitos, 33, 1) = "0" Then
'mnuCadastros.Enabled = False
'End If

'If Mid(StringDireitos, 34, 1) = "0" Then
'mnuFormulario.Enabled = False
'End If

'If Mid(StringDireitos, 37, 1) = "0" Then
'mnuEmissoes.Enabled = False
'End If

Call AtualizaStatusTabelas

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Leave = False Then
    'MsgBox "Não é pertimido sair por aqui. Saia pelas indicações de Saída...", vbCritical, "Saída não permitida..."
    'Cancel = True
    'Exit Sub
    End If
End Sub

Private Sub mnu_CadTabPrecoALTERACAO_Click()
mnuArquivo.Enabled = False
mnuCadastros.Enabled = False
mnuEmissoes.Enabled = False
mnuRelat.Enabled = False
mnuSair.Enabled = False
frmCadTabPrecoALTERACAO_PERC.Show
End Sub

Private Sub mnu_acomp_Click()
frmAIRAcompanha.Show 1
End Sub

Private Sub mnu_CadTabPrecoALTERACAO_UNIT_Click()
mnuArquivo.Enabled = False
mnuCadastros.Enabled = False
mnuEmissoes.Enabled = False
mnuRelat.Enabled = False
mnuSair.Enabled = False
frmCadTabPrecoALTERACAO_UNIT.Show
End Sub

Private Sub mnu_CadTabPrecoINCLUSAO_Click()
mnuArquivo.Enabled = False
mnuCadastros.Enabled = False
mnuEmissoes.Enabled = False
mnuRelat.Enabled = False
mnuSair.Enabled = False
    
frmCadTabPrecoINCLUSAO.Show
End Sub

Private Sub mnu_consultaAWB_Click()
frmConsultaAWB.Show 1
End Sub

Private Sub mnu_voo_Click()
frmAcompAWB.Show 1
End Sub

Private Sub mnuAirAcompanhaREL_Click()
frmAIRRel.Show 1
End Sub

Private Sub mnuAwb_Click()
frmEmissao.Show 1
End Sub

Private Sub mnuCiaAerea_Click()
    frmCadCiaAerea.Show 1
End Sub

Private Sub mnuFormulario_Click()
mnuArquivo.Enabled = False
mnuCadastros.Enabled = False
mnuEmissoes.Enabled = False
mnuRelat.Enabled = False
mnuSair.Enabled = False
    
    frmCadFormulario.Show 1

mnuArquivo.Enabled = True
mnuCadastros.Enabled = True
mnuEmissoes.Enabled = True
mnuRelat.Enabled = True
mnuSair.Enabled = True
End Sub

Private Sub mnuLocal_Click()
mnuArquivo.Enabled = False
mnuCadastros.Enabled = False
mnuEmissoes.Enabled = False
mnuRelat.Enabled = False
mnuSair.Enabled = False

frmCadLocalidade.Show 1

mnuArquivo.Enabled = True
mnuCadastros.Enabled = True
mnuEmissoes.Enabled = True
mnuRelat.Enabled = True
mnuSair.Enabled = True

End Sub

Private Sub mnuProdutos_Click()

End Sub

Private Sub mnuCadProdIATA_Click()
frmCadProdIATA.Show 1
End Sub

Private Sub mnuLote_Click()
frmLote.Show 1
End Sub

Private Sub mnuManifesto_Click()
frmManifesto.Show 1
End Sub

Private Sub mnuPrinters_Click()
frmControleImpressoras.Show 1
End Sub

Private Sub mnuProdutosINT_Click()
    frmCadProdINT.Show 1
End Sub

Private Sub mnuReportCia_Click()
frmReportCIA.Show 1
End Sub

Private Sub mnuRepres_Click()
frmCadRepres.Show 1
End Sub

Private Sub mnuSair_Click()
Leave = True
Unload Me
End Sub


Private Sub mnuVolume_Click()
frmVolLabelTEMP.Show 1
End Sub
