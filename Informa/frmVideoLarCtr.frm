VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVideoLarCtr 
   Caption         =   "Controle Específico VideoLar / Fox Film"
   ClientHeight    =   8175
   ClientLeft      =   795
   ClientTop       =   1350
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin VB.Frame fraRel1 
      Caption         =   "Relatório (Arq. Base)"
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
      Left            =   7200
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "acerto basecli (fantasia, grupo, tipo e cgc)"
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkQuantidades 
         Caption         =   "Processa Quantidades"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optVideolar 
         Caption         =   "Arquivo Videolar (Notas)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optEstudio 
         Caption         =   "Arquivo Cliente Estúdio"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame fraComandos 
      Height          =   5775
      Left            =   10440
      TabIndex        =   29
      Top             =   2160
      Width           =   1575
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   120
         Picture         =   "frmVideoLarCtr.frx":0000
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   1320
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   945
         Left            =   120
         Picture         =   "frmVideoLarCtr.frx":1122
         Stretch         =   -1  'True
         Top             =   2640
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Série NF-Produtora / Cliente Destino / Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   11895
      Begin VB.ComboBox cmbEstudios 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Text            =   "Todos"
         Top             =   330
         Width           =   1575
      End
      Begin VB.CheckBox chkFox 
         Caption         =   "FOX"
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Frame FraFiltroCliente 
         Caption         =   "Filtro de Clientes"
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
         TabIndex        =   41
         Top             =   720
         Width           =   11655
         Begin VB.Frame Frame2 
            Height          =   855
            Left            =   5520
            TabIndex        =   44
            Top             =   120
            Visible         =   0   'False
            Width           =   5535
            Begin VB.CheckBox chkNao 
               Caption         =   "Não"
               Height          =   195
               Left            =   1200
               TabIndex        =   50
               Top             =   200
               Width           =   615
            End
            Begin VB.CommandButton cmdAtualizar 
               Caption         =   "Atualizar"
               Height          =   495
               Left            =   90
               TabIndex        =   12
               Top             =   240
               Width           =   930
            End
            Begin VB.ComboBox cmbGrupo 
               Height          =   315
               Left            =   1080
               TabIndex        =   13
               Top             =   440
               Width           =   2175
            End
            Begin VB.ComboBox cmbTipo 
               Height          =   315
               ItemData        =   "frmVideoLarCtr.frx":44F4
               Left            =   3360
               List            =   "frmVideoLarCtr.frx":44F6
               TabIndex        =   14
               Top             =   440
               Width           =   2055
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Grupo:"
               Height          =   195
               Left            =   2640
               TabIndex        =   46
               Top             =   165
               Width           =   480
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   4920
               TabIndex        =   45
               Top             =   165
               Width           =   360
            End
         End
         Begin VB.Frame frame1 
            Height          =   855
            Left            =   5520
            TabIndex        =   42
            Top             =   120
            Visible         =   0   'False
            Width           =   3975
            Begin VB.TextBox txtFiltroCliente 
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               MaxLength       =   20
               TabIndex        =   20
               Top             =   480
               Width           =   2535
            End
            Begin VB.CommandButton cmdBuscaDest 
               Caption         =   "Busca ..."
               Height          =   255
               Left            =   2880
               TabIndex        =   21
               Top             =   480
               Width           =   855
            End
            Begin VB.Label lblfiltrocliente 
               AutoSize        =   -1  'True
               Caption         =   "Núm. CNPJ:"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   870
            End
         End
         Begin VB.OptionButton optCliTipoGrupo 
            Caption         =   "Por Tipo/Grupo"
            Height          =   195
            Left            =   3360
            TabIndex        =   11
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton optCliFantasia 
            Caption         =   "Por Nome Fantasia"
            Height          =   195
            Left            =   3360
            TabIndex        =   10
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton optCliTodos 
            Caption         =   "Todos Clientes"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optCliCodSap 
            Caption         =   "Por Código SAP"
            Height          =   195
            Left            =   1680
            TabIndex        =   9
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optCliCnpj 
            Caption         =   "Por CNPJ"
            Height          =   195
            Left            =   1680
            TabIndex        =   8
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.TextBox txtPacote 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   10080
         MaxLength       =   30
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtMaterial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   7560
         MaxLength       =   30
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "1"
         Top             =   330
         Width           =   255
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   5640
         TabIndex        =   4
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
         Left            =   4320
         TabIndex        =   3
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
      Begin VB.Label lblPacote 
         AutoSize        =   -1  'True
         Caption         =   "Pacote:"
         Height          =   195
         Left            =   9480
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblMaterial 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Left            =   6870
         TabIndex        =   39
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Série NF:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   3675
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   5520
         TabIndex        =   27
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Resultado da Pesquisa: Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   10215
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1935
         Left            =   1920
         TabIndex        =   48
         Top             =   2040
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   $"frmVideoLarCtr.frx":44F8
            Height          =   1095
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   3375
         End
      End
      Begin TabDlg.SSTab tabVideolar 
         Height          =   4935
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8705
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Notas Fiscais"
         TabPicture(0)   =   "frmVideoLarCtr.frx":45BD
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblTotUnid"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblQtdeNfs"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "flexVideolar"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdGeraArqNfs"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "NFs Produtos"
         TabPicture(1)   =   "frmVideoLarCtr.frx":45D9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flexNFsProds"
         Tab(1).Control(1)=   "flexNFs"
         Tab(1).ControlCount=   2
         Begin VB.CommandButton cmdGeraArqNfs 
            Caption         =   "Gerar Arquivo Notas Fiscais..."
            Height          =   255
            Left            =   6840
            TabIndex        =   17
            Top             =   4560
            Width           =   2895
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexVideolar 
            Height          =   4095
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   7223
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexNFs 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   33
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   7223
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexNFsProds 
            Height          =   4095
            Left            =   -73440
            TabIndex        =   34
            Top             =   360
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   7223
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Qtde. de Notas:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   4560
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Qtde. de Unidades:"
            Height          =   195
            Left            =   2640
            TabIndex        =   37
            Top             =   4560
            Width           =   1380
         End
         Begin VB.Label lblQtdeNfs 
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
            Left            =   1440
            TabIndex        =   36
            Top             =   4560
            Width           =   855
         End
         Begin VB.Label lblTotUnid 
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
            Left            =   4200
            TabIndex        =   35
            Top             =   4560
            Width           =   975
         End
      End
      Begin MSComctlLib.ProgressBar progress 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Enabled         =   0   'False
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   0
      Top             =   2520
   End
End
Attribute VB_Name = "frmVideoLarCtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFox_Click()
    If chkFox.Value = 1 Then
        txtPacote.Visible = True
        lblPacote.Visible = True
        Frame2.Visible = True
        optCliCodSap.Visible = True
        optCliFantasia.Visible = True
        optCliTipoGrupo.Visible = True
        Image1.Visible = True
        cmbEstudios.Visible = False
        optCliCnpj.Visible = True
        tabVideolar.TabEnabled(1) = True
        tabVideolar.Tab = 0
    Else
        txtPacote.Visible = False
        lblPacote.Visible = False
        Frame2.Visible = False
        optCliCodSap.Visible = False
        optCliFantasia.Visible = False
        optCliTipoGrupo.Visible = False
        Image1.Visible = False
        cmbEstudios.Visible = True
        optCliCnpj.Visible = False
        tabVideolar.TabEnabled(1) = False
        tabVideolar.Tab = 0
    End If
End Sub

Private Sub chkFox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmbEstudios_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmbGrupo_Click()
    
    Me.MousePointer = 11
    DoEvents
    DoEvents
    
    'preenche o combo
    If de_informa.rsSel_VLBuscaTipoCli.State = 1 Then de_informa.rsSel_VLBuscaTipoCli.Close
    de_informa.Sel_VLBuscaTipoCli cmbGrupo.List(cmbGrupo.ListIndex)
    
    cmbTipo.Clear
    
    Do Until de_informa.rsSel_VLBuscaTipoCli.EOF
        cmbTipo.AddItem de_informa.rsSel_VLBuscaTipoCli.Fields("tipo")
        de_informa.rsSel_VLBuscaTipoCli.MoveNext
    Loop
    
    Me.MousePointer = 0
    DoEvents
    DoEvents
    
End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdAtualizar_Click()
    'preenche combo de grupo
    
    Me.MousePointer = 11
    DoEvents
    DoEvents
    
    If de_informa.rsSel_VLBuscaGrupoCli.State = 1 Then de_informa.rsSel_VLBuscaGrupoCli.Close
    de_informa.Sel_VLBuscaGrupoCli
    
    cmbGrupo.Clear
    cmbTipo.Clear
    
    Do Until de_informa.rsSel_VLBuscaGrupoCli.EOF
        cmbGrupo.AddItem de_informa.rsSel_VLBuscaGrupoCli.Fields("grupo")
        de_informa.rsSel_VLBuscaGrupoCli.MoveNext
    Loop
    
    Me.MousePointer = 0
    DoEvents
    DoEvents
    
End Sub

Private Sub cmdBuscaDest_Click()
    frmVLBuscaCli.Show 1
    DoEvents
    DoEvents
End Sub

Private Sub cmdGeraArqNfs_Click()
    Dim xcont As Long, xfiles As String, xlinha As String
    If flexVideolar.Rows < 3 Then
        MsgBox "Não há Dados Para Geração de Arquivos ! Primeiro Selecione o Período e clique em Processar.", vbExclamation, "Ops"
    Else
        
        If optEstudio.Value = True Then
            xfiles = "C:\FOX" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & zeros(Hour(Time()), 2) & zeros(Minute(Time()), 2) & ".txt"
        ElseIf optVideolar.Value = True Then
            xfiles = "C:\VL" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & zeros(Hour(Time()), 2) & zeros(Minute(Time()), 2) & ".txt"
        End If
        
        Open xfiles For Output As #1
    
        For xcont = 0 To flexVideolar.Rows - 1
        
            If optEstudio.Value = True Then
            
                xlinha = flexVideolar.TextMatrix(xcont, 1) & "#" & _
                flexVideolar.TextMatrix(xcont, 2) & "#" & _
                flexVideolar.TextMatrix(xcont, 3) & "#" & _
                flexVideolar.TextMatrix(xcont, 4) & "#" & _
                flexVideolar.TextMatrix(xcont, 5) & "#" & _
                flexVideolar.TextMatrix(xcont, 6) & "#" & _
                flexVideolar.TextMatrix(xcont, 7) & "#" & _
                flexVideolar.TextMatrix(xcont, 8) & "#" & _
                flexVideolar.TextMatrix(xcont, 9) & "#" & _
                flexVideolar.TextMatrix(xcont, 10) & "#" & _
                flexVideolar.TextMatrix(xcont, 11) & "#" & _
                flexVideolar.TextMatrix(xcont, 12) & "#" & _
                flexVideolar.TextMatrix(xcont, 13) & "#" & _
                flexVideolar.TextMatrix(xcont, 14) & "#" & _
                flexVideolar.TextMatrix(xcont, 15) & "#" & _
                flexVideolar.TextMatrix(xcont, 16) & "#" & _
                flexVideolar.TextMatrix(xcont, 17) & "#" & _
                flexVideolar.TextMatrix(xcont, 18) & "#" & _
                flexVideolar.TextMatrix(xcont, 19) & "#" & _
                flexVideolar.TextMatrix(xcont, 20) & "#" & _
                flexVideolar.TextMatrix(xcont, 21) & "#" & _
                flexVideolar.TextMatrix(xcont, 22) & "#"
    
            ElseIf optVideolar.Value = True Then
            
                xlinha = flexVideolar.TextMatrix(xcont, 1) & "#" & _
                flexVideolar.TextMatrix(xcont, 2) & "#" & _
                flexVideolar.TextMatrix(xcont, 3) & "#" & _
                flexVideolar.TextMatrix(xcont, 4) & "#" & _
                flexVideolar.TextMatrix(xcont, 5) & "#" & _
                flexVideolar.TextMatrix(xcont, 6) & "#" & _
                flexVideolar.TextMatrix(xcont, 7) & "#" & _
                flexVideolar.TextMatrix(xcont, 8) & "#" & _
                flexVideolar.TextMatrix(xcont, 9) & "#" & _
                flexVideolar.TextMatrix(xcont, 10) & "#" & _
                flexVideolar.TextMatrix(xcont, 11) & "#" & _
                flexVideolar.TextMatrix(xcont, 12) & "#" & _
                flexVideolar.TextMatrix(xcont, 13) & "#" & _
                flexVideolar.TextMatrix(xcont, 14) & "#" & _
                flexVideolar.TextMatrix(xcont, 15) & "#" & _
                flexVideolar.TextMatrix(xcont, 16) & "#"
            
            End If
        
        
            Print #1, xlinha
            DoEvents
        Next
        
        Close #1
        
        MsgBox "OK ! Arquivo Gerado em " & xfiles & "." & Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
        "O Arquivo Gerado é do Tipo Texto ( TXT com Delimitador # ) e você pode abrí-lo em diversos aplicativos. Para Abrí-lo no MS-Excel, em ABRIR escolha ARQUIVOS DO TIPO = Arquivos de Texto e selecione o arquivo no local indicado acima. Na Caixa ASSISTENTE DE IMPORTAÇÃO escolha DELIMITADO e o caracter delimitador escolha OUTROS e digite # . Clique em Concluir e o arquivo será importado para o MS-Excel.", vbInformation, "Geração de Arquivo TXT"

    End If
        

End Sub

Private Sub CmdProcessar_Click()
    Dim xdataper1 As Date, xdataper2 As Date, xocorr As String, xLin As Long, xtotqtde As Long
    Dim xcgccli As String, xCodSap As String, xFantasia As String, xGrupo As String, xTipo As String
    Dim xrs As Recordset, xStatus As String
    
    
    If IsDate(mskPer1) And IsDate(mskPer2) Then
        If CDate(mskPer1) > CDate(mskPer2) Then
            MsgBox "Período Escolhido Inválido !", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
'        If Abs(CDate(mskPer2) - CDate(mskPer1)) > 15 Then
'            MsgBox "Período Escolhido Maior que 15 dias. Escolha um período menor.", vbCritical, "Erro"
'            mskPer1.SetFocus
'             Exit Sub
'        End If
    Else
        MsgBox "Datas de Período Inválido !", vbCritical, "Erro"
        mskPer1.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Trim$(txtSerie)) Then
        MsgBox "Série Inválida !"
        txtSerie.SetFocus
        Exit Sub
    End If
    
    flexVideolar.Rows = 1
    flexVideolar.Rows = 2
    flexVideolar.FixedRows = 1
    flexNFs.Rows = 1
    flexNFs.Rows = 2
    flexNFs.FixedRows = 1
    flexNFsProds.Rows = 1
    flexNFsProds.Rows = 2
    flexNFsProds.FixedRows = 1
    
    xdataper1 = CDate(mskPer1)
    xdataper2 = CDate(mskPer2)
    
    If chkFox.Value = 1 Then
    
        If optCliTodos.Value = True Then
            xcgccli = "%"
            xCodSap = "%"
            xFantasia = "%"
            xGrupo = "%"
            xTipo = "%"
        ElseIf optCliCnpj.Value = True Then
            If Len(Trim$(txtFiltroCliente)) < 8 Then
                MsgBox "Número de CNPJ Menor que 8 Números. Inválido !"
                txtFiltroCliente.SetFocus
                Exit Sub
            End If
            xcgccli = Trim$(txtFiltroCliente) & "%"
            xCodSap = "%"
            xFantasia = "%"
            xGrupo = "%"
            xTipo = "%"
        ElseIf optCliCodSap.Value = True Then
            If Len(Trim$(txtFiltroCliente)) < 9 Then
                MsgBox "Número de Código SAP Menor que 9 Números. Inválido !"
                txtFiltroCliente.SetFocus
                Exit Sub
            End If
            xcgccli = "%"
            xCodSap = Trim$(txtFiltroCliente) & "%"
            xFantasia = "%"
            xGrupo = "%"
            xTipo = "%"
        ElseIf optCliFantasia.Value = True Then
            If Len(Trim$(txtFiltroCliente)) < 2 Then
                MsgBox "Nome Fantasia Inválido !"
                txtFiltroCliente.SetFocus
                Exit Sub
            End If
            xcgccli = "%"
            xCodSap = "%"
            xFantasia = Trim$(txtFiltroCliente)
            xGrupo = "%"
            xTipo = "%"
        ElseIf optCliTipoGrupo.Value = True Then
            If Len(Trim$(cmbGrupo.List(cmbGrupo.ListIndex))) < 1 And Len(Trim$(cmbTipo.List(cmbTipo.ListIndex))) < 1 Then
                MsgBox "Você Deve Escolhe o Grupo ou Tipo do Cliente !"
                optCliTipoGrupo.SetFocus
                Exit Sub
            End If
            xcgccli = "%"
            xCodSap = "%"
            xFantasia = "%"
            If Len(Trim$(cmbGrupo.List(cmbGrupo.ListIndex))) < 1 Then
                xGrupo = "%"
            Else
                xGrupo = cmbGrupo.List(cmbGrupo.ListIndex)
            End If
            If Len(Trim$(cmbTipo.List(cmbTipo.ListIndex))) < 1 Then
                xTipo = "%"
            Else
                xTipo = cmbTipo.List(cmbTipo.ListIndex)
            End If
        End If
        
        Me.MousePointer = 11
        fraCliente.Enabled = False
        fraGrid.Enabled = False
        fraGrid.Caption = "Resultado da Pesquisa: AGUARDE ..."
        fraComandos.Enabled = False
        fraRel1.Enabled = False
        DoEvents
        
        If chkNao.Visible = True And chkNao.Value = 1 Then
            If de_informa.rsSel_BuscaNFBaseCliNAO.State = 1 Then de_informa.rsSel_BuscaNFBaseCliNAO.Close
            de_informa.Sel_BuscaNFBaseCliNAO Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), _
                                          "%" & Trim$(txtMaterial) & "%", "%" & Trim$(txtPacote) & "%", _
                                          xcgccli, xCodSap, xFantasia, xGrupo, xTipo
            Set xrs = de_informa.rsSel_BuscaNFBaseCliNAO
        Else
            If de_informa.rsSel_BuscaNFBaseCli.State = 1 Then de_informa.rsSel_BuscaNFBaseCli.Close
            de_informa.Sel_BuscaNFBaseCli Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), _
                                          "%" & Trim$(txtMaterial) & "%", "%" & Trim$(txtPacote) & "%", _
                                          xcgccli, xCodSap, xFantasia, xGrupo, xTipo
            Set xrs = de_informa.rsSel_BuscaNFBaseCli
        End If
        
        If xrs.RecordCount < 1 Then
            MsgBox "Não Há Dados Para as Opções Selecionadas !"
            Me.MousePointer = 0
            fraCliente.Enabled = True
            fraGrid.Enabled = True
            fraGrid.Caption = "Resultado da Pesquisa:"
            fraComandos.Enabled = True
            fraRel1.Enabled = True
            DoEvents
            If txtSerie.Enabled = True Then
                txtSerie.SetFocus
            Else
                mskPer1.SetFocus
            End If
            Exit Sub
        End If
        
        xLin = 1
        flexVideolar.Rows = xrs.RecordCount + 1
        flexNFs.Rows = xrs.RecordCount + 1
        
        progress.Max = xrs.RecordCount
        xtotqtde = 0
        DoEvents
        
        Do Until xrs.EOF
        
            flexVideolar.Row = xLin
            flexNFs.Row = xLin
            progress.Value = xLin
            DoEvents
            
            'preenche os dados referente a tabela base do cliente (basecli)
            flexVideolar.Col = 1
            flexVideolar.Text = xrs.Fields("numnfnum")
            flexVideolar.Col = 2
            flexVideolar.Text = Year(xrs.Fields("datanf")) & "/" & _
                                zeros(Month(xrs.Fields("datanf")), 2) & "/" & _
                                zeros(Day(xrs.Fields("datanf")), 2)
            flexVideolar.Col = 3
            flexVideolar.Text = xrs.Fields("clientenf")
            flexVideolar.Col = 4
            flexVideolar.Text = xrs.Fields("cidadenf")
            flexVideolar.Col = 5
            flexVideolar.Text = xrs.Fields("ufnf")
            flexVideolar.Col = 6
            flexVideolar.Text = xrs.Fields("entr_solic")
            flexVideolar.Col = 7
            flexVideolar.Text = xrs.Fields("pacote")
            flexVideolar.Col = 8
            flexVideolar.Text = xrs.Fields("grupoclinf")
            flexVideolar.Col = 9
            flexVideolar.Text = xrs.Fields("tipoclinf")
            flexVideolar.Col = 10
            flexVideolar.Text = xrs.Fields("itens")
            
            xtotqtde = xtotqtde + CDbl(xrs.Fields("itens"))
            
            flexNFs.Col = 1
            
            flexNFs.Text = xrs.Fields("numnfnum")
            
            flexVideolar.Col = 11
            If IsNull(xrs.Fields("datavideolar")) Then
                flexVideolar.Text = ""
            Else
                flexVideolar.Text = Year(xrs.Fields("datavideolar")) & "/" & _
                                    zeros(Month(xrs.Fields("datavideolar")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("datavideolar")), 2) & " " & _
                                    zeros(Hour(xrs.Fields("datavideolar")), 2) & ":" & _
                                    zeros(Minute(xrs.Fields("datavideolar")), 2)
            End If
            flexVideolar.Col = 12
            If IsNull(xrs.Fields("datafluxo")) Then
                flexVideolar.Text = ""
            Else
                flexVideolar.Text = Year(xrs.Fields("datafluxo")) & "/" & _
                                    zeros(Month(xrs.Fields("datafluxo")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("datafluxo")), 2) & " " & _
                                    zeros(Hour(xrs.Fields("datafluxo")), 2) & ":" & _
                                    zeros(Minute(xrs.Fields("datafluxo")), 2)
            End If

            If IsNull(xrs.Fields("filialctc")) Then
                flexVideolar.Col = 13
                flexVideolar.Text = ""
                flexVideolar.Col = 14
                flexVideolar.Text = ""
                flexVideolar.Col = 15
                flexVideolar.Text = ""
                flexVideolar.Col = 16
                flexVideolar.Text = ""
                flexVideolar.Col = 22
                flexVideolar.Text = ""
            Else
                flexVideolar.Col = 13
                flexVideolar.Text = xrs.Fields("filialctc")
                flexVideolar.Col = 14
                flexVideolar.Text = Year(xrs.Fields("datactr")) & "/" & _
                                    zeros(Month(xrs.Fields("datactr")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("datactr")), 2)
                flexVideolar.Col = 15
                flexVideolar.Text = xrs.Fields("modal")
                flexVideolar.Col = 16
                flexVideolar.Text = Year(xrs.Fields("prev_entrega")) & "/" & _
                                    zeros(Month(xrs.Fields("prev_entrega")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("prev_entrega")), 2)
                flexVideolar.Col = 19
                flexVideolar.Text = xrs.Fields("transp_sub")
                flexVideolar.Col = 23
                flexVideolar.Text = xrs.Fields("obs_emissao")
            End If
            
            DoEvents
            
            flexVideolar.Col = 17
            flexVideolar.Text = ""
            flexVideolar.Col = 18
            flexVideolar.Text = ""
            flexVideolar.Col = 20
            flexVideolar.Text = ""
            flexVideolar.Col = 21
            flexVideolar.Text = ""
            flexVideolar.Col = 23
            flexVideolar.Text = ""
                
            If Len(xrs.Fields("filialctc")) > 8 Then
           
                'busca dados de MANIFESTO
                If de_informa.rsSel_ManifestoPorCTC.State = 1 Then de_informa.rsSel_ManifestoPorCTC.Close
                de_informa.Sel_ManifestoPorCTC xrs.Fields("filialctc")
                
                If de_informa.rsSel_ManifestoPorCTC.RecordCount > 0 Then
                    de_informa.rsSel_ManifestoPorCTC.MoveLast
                    
                    flexVideolar.Col = 17
                    flexVideolar.Text = de_informa.rsSel_ManifestoPorCTC.Fields("filialmanifesto")
                    flexVideolar.Col = 18
                    flexVideolar.Text = Year(de_informa.rsSel_ManifestoPorCTC.Fields("dtemissao")) & "/" & _
                                        zeros(Month(de_informa.rsSel_ManifestoPorCTC.Fields("dtemissao")), 2) & "/" & _
                                        zeros(Day(de_informa.rsSel_ManifestoPorCTC.Fields("dtemissao")), 2)
                End If
           
                'busca dados de ocorrência ou entrega
                If de_informa.rsSel_OcorrenciasCTR.State = 1 Then de_informa.rsSel_OcorrenciasCTR.Close
                de_informa.Sel_OcorrenciasCTR xrs.Fields("filialctc")
                
                xocorr = ""
                
                xStatus = "SEM POSICAO"
                
                Do Until de_informa.rsSel_OcorrenciasCTR.EOF
                    If de_informa.rsSel_OcorrenciasCTR.Fields("cod_ocorr") = "01" Then
                        flexVideolar.Col = 21
                        flexVideolar.Text = Year(de_informa.rsSel_OcorrenciasCTR.Fields("data")) & "/" & _
                                    zeros(Month(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2) & "/" & _
                                    zeros(Day(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2)
                                    
                        xStatus = "ENTREGUE"
                        
                    Else
                        If de_informa.rsSel_OcorrenciasCTR.Fields("cod_ocorr") = "00" Then
                            xStatus = "BX.SEM ENTREGA"
                        End If
                        If xStatus = "SEM POSICAO" Then
                            xStatus = "EM OCORRENCIA"
                        End If
                        xocorr = CVar(zeros(Day(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2)) & "/" & _
                                 CVar(zeros(Month(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2)) & "/" & _
                                 Trim$(CVar(Year(de_informa.rsSel_OcorrenciasCTR.Fields("data")))) & "-" & _
                                 Trim$(de_informa.rsSel_OcorrenciasCTR.Fields("descr_ocorr"))
                                 
                        If Not IsNull(de_informa.rsSel_OcorrenciasCTR.Fields("obs_ocorr")) Then
                            If Len(Trim$(de_informa.rsSel_OcorrenciasCTR.Fields("obs_ocorr"))) > 0 Then
                                xocorr = xocorr & " (" & Trim$(LCase(de_informa.rsSel_OcorrenciasCTR.Fields("obs_ocorr"))) & ")"
                            End If
                        End If
                        xocorr = xocorr & " ; "
                    End If
                    de_informa.rsSel_OcorrenciasCTR.MoveNext
                Loop
                
                flexVideolar.Col = 22
                flexVideolar.Text = xocorr
                
                If xStatus = "SEM POSICAO" Then
                    If xrs.Fields("prev_entrega") >= datahora("data") Then
                        xStatus = "EM TRANSITO"
                    Else
                        xStatus = "SEM POSICAO HA " & Trim$(Str(Val(datahora("data") - CDate(Mid$(CVar(xrs.Fields("prev_entrega")), 1, 10))))) & " dia(s)"
                    End If
                End If
                
                flexVideolar.Col = 20
                flexVideolar.Text = xStatus
                
            End If
            
            DoEvents
            
            xrs.MoveNext
            
            xLin = xLin + 1
            
        Loop
        
        If chkNao.Visible = True And chkNao.Value = 1 Then
                
            'busca totais de NFS
            If de_informa.rsSel_VLTotNFSNao.State = 1 Then de_informa.rsSel_VLTotNFSNao.Close
            de_informa.Sel_VLTotNFSNao Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), _
                                          "%" & Trim$(txtMaterial) & "%", "%" & Trim$(txtPacote) & "%", _
                                          xcgccli, xCodSap, xFantasia, xGrupo, xTipo
            
            'busca totais de itens
            If de_informa.rsSel_VLTotItensNao.State = 1 Then de_informa.rsSel_VLTotItensNao.Close
            de_informa.Sel_VLTotItensNao Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), _
                                          "%" & Trim$(txtMaterial) & "%", "%" & Trim$(txtPacote) & "%", _
                                          xcgccli, xCodSap, xFantasia, xGrupo, xTipo
        
            lblQtdeNfs = Format(de_informa.rsSel_VLTotNFSNao.RecordCount, "###,##0")
            lblTotUnid = Format(de_informa.rsSel_VLTotItensNao.Fields("qtde"), "###,##0")
            
        Else
        
            'busca totais de NFS
            If de_informa.rsSel_VLTotNFS.State = 1 Then de_informa.rsSel_VLTotNFS.Close
            de_informa.Sel_VLTotNFS Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), _
                                          "%" & Trim$(txtMaterial) & "%", "%" & Trim$(txtPacote) & "%", _
                                          xcgccli, xCodSap, xFantasia, xGrupo, xTipo
            
            'busca totais de itens
            If de_informa.rsSel_VLTotItens.State = 1 Then de_informa.rsSel_VLTotItens.Close
            de_informa.Sel_VLTotItens Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), _
                                          "%" & Trim$(txtMaterial) & "%", "%" & Trim$(txtPacote) & "%", _
                                          xcgccli, xCodSap, xFantasia, xGrupo, xTipo
        
            lblQtdeNfs = Format(de_informa.rsSel_VLTotNFS.RecordCount, "###,##0")
            lblTotUnid = Format(de_informa.rsSel_VLTotItens.Fields("qtde"), "###,##0")
        
        End If
        
        Me.MousePointer = 0
        
        MsgBox "Final de Processamento dos Dados. Para Gerar Arquivo Clique no Botão Abaixo <Gerar Arquivo Notas Fiscais>."
        

    ElseIf chkFox.Value = 0 Then
    
        Me.MousePointer = 11
        fraCliente.Enabled = False
        fraGrid.Enabled = False
        fraGrid.Caption = "Resultado da Pesquisa: AGUARDE ..."
        fraComandos.Enabled = False
        fraRel1.Enabled = False
        DoEvents
        
        If cmbEstudios.Text = "Todos" Then
            xcodlocal = "%"
        Else
            xcodlocal = Mid$(cmbEstudios.Text, 1, 6)
        End If
        
        If de_informa.rsSel_BaseCliCAPs.State = 1 Then de_informa.rsSel_BaseCliCAPs.Close
        de_informa.Sel_BaseCliCAPs Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), xcodlocal, Trim$(txtMaterial) & "%"
        
        Set xrs = de_informa.rsSel_BaseCliCAPs
        
        If xrs.RecordCount < 1 Then
            MsgBox "Não Há Dados Para as Opções Selecionadas !"
            Me.MousePointer = 0
            fraCliente.Enabled = True
            fraGrid.Enabled = True
            fraGrid.Caption = "Resultado da Pesquisa:"
            fraComandos.Enabled = True
            fraRel1.Enabled = True
            DoEvents
            If txtSerie.Enabled = True Then
                txtSerie.SetFocus
            Else
                mskPer1.SetFocus
            End If
            Exit Sub
        End If
        
        xLin = 1
        flexVideolar.Rows = xrs.RecordCount + 1
        flexNFs.Rows = xrs.RecordCount + 1
        
        progress.Max = xrs.RecordCount
        xtotqtde = 0
        DoEvents
        
        Do Until xrs.EOF
        
            flexVideolar.Row = xLin
            flexNFs.Row = xLin
            progress.Value = xLin
            DoEvents
            
            'preenche os dados referente a tabela base do cliente (basecli)
            flexVideolar.Col = 1
            flexVideolar.Text = xrs.Fields("numnfnum")
            flexVideolar.Col = 2
            flexVideolar.Text = Year(xrs.Fields("emissaonf")) & "/" & _
                                zeros(Month(xrs.Fields("emissaonf")), 2) & "/" & _
                                zeros(Day(xrs.Fields("emissaonf")), 2)
            flexVideolar.Col = 3
            flexVideolar.Text = xrs.Fields("dest_nome")
            flexVideolar.Col = 4
            flexVideolar.Text = xrs.Fields("dest_cidade")
            flexVideolar.Col = 5
            flexVideolar.Text = xrs.Fields("dest_uf")
            flexVideolar.Col = 6
            flexVideolar.Text = xrs.Fields("entregasolic")
            flexVideolar.Col = 7
            flexVideolar.Text = ""
            flexVideolar.Col = 8
            flexVideolar.Text = ""
            flexVideolar.Col = 9
            flexVideolar.Text = xrs.Fields("codlocal")
            flexVideolar.Col = 10
            flexVideolar.Text = xrs.Fields("qtde")
            
            flexNFs.Col = 1
            
            flexNFs.Text = xrs.Fields("numnfnum")
            
            flexVideolar.Col = 11
            If IsNull(xrs.Fields("datainterface")) Then
                flexVideolar.Text = ""
            Else
                flexVideolar.Text = Year(xrs.Fields("datainterface")) & "/" & _
                                    zeros(Month(xrs.Fields("datainterface")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("datainterface")), 2) & " " & _
                                    zeros(Hour(xrs.Fields("datainterface")), 2) & ":" & _
                                    zeros(Minute(xrs.Fields("datainterface")), 2)
            End If
            flexVideolar.Col = 12
            If IsNull(xrs.Fields("receb_luft")) Then
                flexVideolar.Text = ""
            Else
                flexVideolar.Text = Year(xrs.Fields("receb_luft")) & "/" & _
                                    zeros(Month(xrs.Fields("receb_luft")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("receb_luft")), 2) & " " & _
                                    zeros(Hour(xrs.Fields("receb_luft")), 2) & ":" & _
                                    zeros(Minute(xrs.Fields("receb_luft")), 2)
            End If

            If IsNull(xrs.Fields("filialctc")) Then
                flexVideolar.Col = 13
                flexVideolar.Text = ""
                flexVideolar.Col = 14
                flexVideolar.Text = ""
                flexVideolar.Col = 15
                flexVideolar.Text = ""
                flexVideolar.Col = 16
                flexVideolar.Text = ""
                flexVideolar.Col = 22
                flexVideolar.Text = ""
            Else
                flexVideolar.Col = 13
                flexVideolar.Text = xrs.Fields("filialctc")
                flexVideolar.Col = 14
                flexVideolar.Text = Year(xrs.Fields("dataCTC")) & "/" & _
                                    zeros(Month(xrs.Fields("dataCTC")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("dataCTC")), 2)
                flexVideolar.Col = 15
                flexVideolar.Text = xrs.Fields("modal")
                flexVideolar.Col = 16
                flexVideolar.Text = Year(xrs.Fields("prev_entrega")) & "/" & _
                                    zeros(Month(xrs.Fields("prev_entrega")), 2) & "/" & _
                                    zeros(Day(xrs.Fields("prev_entrega")), 2)
                flexVideolar.Col = 19
                flexVideolar.Text = xrs.Fields("transp_sub")
                flexVideolar.Col = 23
                flexVideolar.Text = xrs.Fields("obs_emissao")
            End If
            
            DoEvents
            
            flexVideolar.Col = 17
            flexVideolar.Text = ""
            flexVideolar.Col = 18
            flexVideolar.Text = ""
            flexVideolar.Col = 20
            flexVideolar.Text = ""
            flexVideolar.Col = 21
            flexVideolar.Text = ""
            flexVideolar.Col = 23
            flexVideolar.Text = ""
                
            If Len(xrs.Fields("filialctc")) > 8 Then
           
                'busca dados de MANIFESTO
                If de_informa.rsSel_ManifestoPorCTC.State = 1 Then de_informa.rsSel_ManifestoPorCTC.Close
                de_informa.Sel_ManifestoPorCTC xrs.Fields("filialctc")
                
                If de_informa.rsSel_ManifestoPorCTC.RecordCount > 0 Then
                    de_informa.rsSel_ManifestoPorCTC.MoveLast
                    
                    flexVideolar.Col = 17
                    flexVideolar.Text = de_informa.rsSel_ManifestoPorCTC.Fields("filialmanifesto")
                    flexVideolar.Col = 18
                    flexVideolar.Text = Year(de_informa.rsSel_ManifestoPorCTC.Fields("dtemissao")) & "/" & _
                                        zeros(Month(de_informa.rsSel_ManifestoPorCTC.Fields("dtemissao")), 2) & "/" & _
                                        zeros(Day(de_informa.rsSel_ManifestoPorCTC.Fields("dtemissao")), 2)
                End If
           
                'busca dados de ocorrência ou entrega
                If de_informa.rsSel_OcorrenciasCTR.State = 1 Then de_informa.rsSel_OcorrenciasCTR.Close
                de_informa.Sel_OcorrenciasCTR xrs.Fields("filialctc")
                
                xocorr = ""
                
                xStatus = "SEM POSICAO"
                
                Do Until de_informa.rsSel_OcorrenciasCTR.EOF
                    If de_informa.rsSel_OcorrenciasCTR.Fields("cod_ocorr") = "01" Then
                        flexVideolar.Col = 21
                        flexVideolar.Text = Year(de_informa.rsSel_OcorrenciasCTR.Fields("data")) & "/" & _
                                    zeros(Month(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2) & "/" & _
                                    zeros(Day(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2)
                                    
                        xStatus = "ENTREGUE"
                        
                    Else
                        If de_informa.rsSel_OcorrenciasCTR.Fields("cod_ocorr") = "00" Then
                            xStatus = "BX.SEM ENTREGA"
                        End If
                        If xStatus = "SEM POSICAO" Then
                            xStatus = "EM OCORRENCIA"
                        End If
                        xocorr = CVar(zeros(Day(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2)) & "/" & _
                                 CVar(zeros(Month(de_informa.rsSel_OcorrenciasCTR.Fields("data")), 2)) & "/" & _
                                 Trim$(CVar(Year(de_informa.rsSel_OcorrenciasCTR.Fields("data")))) & "-" & _
                                 Trim$(de_informa.rsSel_OcorrenciasCTR.Fields("descr_ocorr"))
                                 
                        If Not IsNull(de_informa.rsSel_OcorrenciasCTR.Fields("obs_ocorr")) Then
                            If Len(Trim$(de_informa.rsSel_OcorrenciasCTR.Fields("obs_ocorr"))) > 0 Then
                                xocorr = xocorr & " (" & Trim$(LCase(de_informa.rsSel_OcorrenciasCTR.Fields("obs_ocorr"))) & ")"
                            End If
                        End If
                        xocorr = xocorr & " ; "
                    End If
                    de_informa.rsSel_OcorrenciasCTR.MoveNext
                Loop
                
                flexVideolar.Col = 22
                flexVideolar.Text = xocorr
                
                If xStatus = "SEM POSICAO" Then
                    If xrs.Fields("prev_entrega") >= datahora("data") Then
                        xStatus = "EM TRANSITO"
                    Else
                        xStatus = "SEM POSICAO HA " & Trim$(Str(Val(datahora("data") - CDate(Mid$(CVar(xrs.Fields("prev_entrega")), 1, 10))))) & " dia(s)"
                    End If
                End If
                
                flexVideolar.Col = 20
                flexVideolar.Text = xStatus
                
            End If
            
            DoEvents
            
            xrs.MoveNext
            
            xLin = xLin + 1
            
        Loop
        
        'busca totais de NFS
        If de_informa.rsSel_VLTotNFSCaps.State = 1 Then de_informa.rsSel_VLTotNFSCaps.Close
        de_informa.Sel_VLTotNFSCaps Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), xcodlocal, Trim$(txtMaterial) & "%"
        
        'busca totais de itens
        If de_informa.rsSel_VLTotItensCaps.State = 1 Then de_informa.rsSel_VLTotItensCaps.Close
        de_informa.Sel_VLTotItensCaps Trim$(txtSerie), CDate(xdataper1), CDate(xdataper2), xcodlocal, Trim$(txtMaterial) & "%"
    
        lblQtdeNfs = Format(de_informa.rsSel_VLTotNFSCaps.RecordCount, "###,##0")
        lblTotUnid = Format(de_informa.rsSel_VLTotItensCaps.Fields("qtde"), "###,##0")
        
        Me.MousePointer = 0
        
        MsgBox "Final de Processamento dos Dados. Para Gerar Arquivo Clique no Botão Abaixo <Gerar Arquivo Notas Fiscais>."
        
    ElseIf optCheck.Value = True Then
    
    
    ElseIf optEmitidos.Value = True Then


    End If

    Me.MousePointer = 0
    fraCliente.Enabled = True
    fraGrid.Enabled = True
    fraComandos.Enabled = True
    fraRel1.Enabled = True

    fraGrid.Caption = "Resultado da Pesquisa: Relatório Por Período (Base Cliente Estúdio)"
    DoEvents
                        
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command3_Click()
    txtDestCGC.Text = ""
    txtMaterial.Text = ""
    txtPacote.Text = ""
    lblDestNome.Caption = ""
End Sub

Private Sub Command1_Click()
    
    If de_informa.rsSel_VLAcerto1.State = 1 Then de_informa.rsSel_VLAcerto1.Close
    de_informa.Sel_VLAcerto1
    
    Do Until de_informa.rsSel_VLAcerto1.EOF
        If de_informa.rsSel_VLClienteSAP.State = 1 Then de_informa.rsSel_VLClienteSAP.Close
        de_informa.Sel_VLClienteSAP de_informa.rsSel_VLAcerto1.Fields("codclinf")
        
        If de_informa.rsSel_VLClienteSAP.RecordCount > 0 Then
            de_informa.Alt_VLAcerto1 de_informa.rsSel_VLClienteSAP.Fields("cliente_cgc"), _
                                      de_informa.rsSel_VLClienteSAP.Fields("cliente_fantasia"), _
                                      de_informa.rsSel_VLClienteSAP.Fields("grupo"), _
                                      de_informa.rsSel_VLClienteSAP.Fields("tipo"), _
                                      de_informa.rsSel_VLAcerto1.Fields("codclinf")
        End If
            
        de_informa.rsSel_VLAcerto1.MoveNext
    
    Loop
                                    
    MsgBox "ok"
    
End Sub

Private Sub Command2_Click()

End Sub

Private Sub flexNFs_Click()
    Dim xLin As Long
    
    flexNFsProds.Rows = 1
    flexNFsProds.Rows = 2
    flexNFsProds.FixedRows = 1
    
    Me.MousePointer = 11
    fraCliente.Enabled = False
    fraGrid.Enabled = False
    fraComandos.Enabled = False
    fraRel1.Enabled = False
    DoEvents
    
    If de_informa.rsSel_NFsProds.State = 1 Then de_informa.rsSel_NFsProds.Close
    de_informa.Sel_NFsProds "04229761000413", flexNFs.TextMatrix(flexNFs.Row, 1), Trim$(txtSerie)
    
    xLin = 1
    flexNFsProds.Rows = de_informa.rsSel_NFsProds.RecordCount + 1
    progress.Max = de_informa.rsSel_NFsProds.RecordCount
    DoEvents
    
    Do Until de_informa.rsSel_NFsProds.EOF
    
        flexNFsProds.Row = xLin
        progress.Value = xLin
        DoEvents
        
        'preenche os dados referente a tabela base do cliente (basecli)
        flexNFsProds.Col = 1
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("ordvenda")
        flexNFsProds.Col = 2
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("pedido")
        flexNFsProds.Col = 3
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("codclinf")
        flexNFsProds.Col = 4
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("clientenf")
        flexNFsProds.Col = 5
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("cidadenf")
        flexNFsProds.Col = 6
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("ufnf")
        flexNFsProds.Col = 7
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("codmaterial")
        flexNFsProds.Col = 8
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("material")
        flexNFsProds.Col = 9
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("qtdeitem")
        flexNFsProds.Col = 10
        flexNFsProds.Text = de_informa.rsSel_NFsProds.Fields("datanf")
        
        DoEvents
        
        de_informa.rsSel_NFsProds.MoveNext
        
        xLin = xLin + 1
        
    Loop
    
    Me.MousePointer = 0
    fraCliente.Enabled = True
    fraGrid.Enabled = True
    fraComandos.Enabled = True
    fraRel1.Enabled = True
    DoEvents
    
End Sub
Private Sub flexNFs_Scroll()
    flexVideolar.TopRow = flexNFs.TopRow
End Sub

Private Sub flexVideolar_Click()
    flexVideolar.ToolTipText = flexVideolar.TextMatrix(flexVideolar.Row, flexVideolar.Col)
    Timer1.Interval = 20000
End Sub
Private Sub flexVideolar_Scroll()
    flexNFs.TopRow = flexVideolar.TopRow
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    
    tabVideolar.TabEnabled(1) = False
    
    If xusuario = "FOXFILM" Then
        chkFox.Value = 1
        chkFox_Click
        chkFox.Enabled = False
    End If
    
    flexVideolar.Cols = 24
    flexVideolar.Row = 0
    
    flexVideolar.Col = 1
    flexVideolar.Text = "NF"
    flexVideolar.Col = 2
    flexVideolar.Text = "Emissão"
    flexVideolar.Col = 3
    flexVideolar.Text = "Cliente"
    flexVideolar.Col = 4
    flexVideolar.Text = "Cidade"
    flexVideolar.Col = 5
    flexVideolar.Text = "UF"
    flexVideolar.Col = 6
    flexVideolar.Text = "Entrega Solic."
    flexVideolar.Col = 7
    flexVideolar.Text = "Pacote"
    flexVideolar.Col = 8
    flexVideolar.Text = "Grupo"
    flexVideolar.Col = 9
    flexVideolar.Text = "Tipo"
    flexVideolar.Col = 10
    flexVideolar.Text = "Qtde."
    flexVideolar.Col = 11
    flexVideolar.Text = "Dt.Arq.Videolar"
    flexVideolar.Col = 12
    flexVideolar.Text = "Dt.Rec.Luft"
    flexVideolar.Col = 13
    flexVideolar.Text = "CTR Intec"
    flexVideolar.Col = 14
    flexVideolar.Text = "Dt.CTR Intec"
    flexVideolar.Col = 15
    flexVideolar.Text = "Modal"
    flexVideolar.Col = 16
    flexVideolar.Text = "Prev.Entrega"
    flexVideolar.Col = 17
    flexVideolar.Text = "Manifesto"
    flexVideolar.Col = 18
    flexVideolar.Text = "Data Manif."
    flexVideolar.Col = 19
    flexVideolar.Text = "Redespacho"
    flexVideolar.Col = 20
    flexVideolar.Text = "Status"
    flexVideolar.Col = 21
    flexVideolar.Text = "Data Entrega"
    flexVideolar.Col = 22
    flexVideolar.Text = "Ocorrências"
    flexVideolar.Col = 23
    flexVideolar.Text = "Obs.Emissão"
'   flexVideolar.Col = 17
'   flexVideolar.Text = "Redespacho"
'   flexVideolar.Col = 18
'   flexVideolar.Text = "Data Entrega"
'   flexVideolar.Col = 19
'   flexVideolar.Text = "Ocorrências"
'   flexVideolar.Col = 20
'   flexVideolar.Text = "Obs.Emissão"
    flexVideolar.ColWidth(0) = 200
    flexVideolar.ColWidth(1) = 750
    flexVideolar.ColWidth(2) = 1100
    flexVideolar.ColWidth(3) = 2500
    flexVideolar.ColWidth(4) = 2000
    flexVideolar.ColWidth(5) = 350
    flexVideolar.ColWidth(6) = 1200
    flexVideolar.ColWidth(7) = 2000
    flexVideolar.ColWidth(8) = 1200
    flexVideolar.ColWidth(9) = 1200
    flexVideolar.ColWidth(10) = 600
    flexVideolar.ColWidth(11) = 1450
    flexVideolar.ColWidth(12) = 1450
    flexVideolar.ColWidth(13) = 1150
    flexVideolar.ColWidth(14) = 1200
    flexVideolar.ColWidth(15) = 1200
    flexVideolar.ColWidth(16) = 1200
    flexVideolar.ColWidth(17) = 1000
    flexVideolar.ColWidth(18) = 1100
    flexVideolar.ColWidth(19) = 2000
    flexVideolar.ColWidth(20) = 2000
    flexVideolar.ColWidth(21) = 1200
    flexVideolar.ColWidth(22) = 9200
    flexVideolar.ColWidth(23) = 9200
    
    flexVideolar.ColAlignment(1) = 1
    flexVideolar.ColAlignment(2) = 1
    flexVideolar.ColAlignment(3) = 1
    flexVideolar.ColAlignment(4) = 1
    flexVideolar.ColAlignment(5) = 1
    flexVideolar.ColAlignment(6) = 1
    flexVideolar.ColAlignment(7) = 1
    flexVideolar.ColAlignment(8) = 1
    flexVideolar.ColAlignment(9) = 1
    flexVideolar.ColAlignment(10) = 1
    flexVideolar.ColAlignment(11) = 1
    flexVideolar.ColAlignment(12) = 1
    flexVideolar.ColAlignment(13) = 1
    flexVideolar.ColAlignment(14) = 1
    flexVideolar.ColAlignment(15) = 1
    flexVideolar.ColAlignment(16) = 1
    flexVideolar.ColAlignment(17) = 1
    flexVideolar.ColAlignment(18) = 1
    flexVideolar.ColAlignment(19) = 1
    flexVideolar.ColAlignment(20) = 1
    flexVideolar.ColAlignment(21) = 1
    flexVideolar.ColAlignment(22) = 1
    flexVideolar.ColAlignment(23) = 1
    
    flexNFs.Cols = 2
    flexNFs.Row = 0
    flexNFs.Col = 1
    flexNFs.Text = "NF"
    flexNFs.ColWidth(0) = 200
    flexNFs.ColWidth(1) = 800
    
    flexNFsProds.Cols = 11
    flexNFsProds.Row = 0
    
    flexNFsProds.Col = 1
    flexNFsProds.Text = "Ord.Venda"
    flexNFsProds.Col = 2
    flexNFsProds.Text = "Pedido"
    flexNFsProds.Col = 3
    flexNFsProds.Text = "Cód. Cliente"
    flexNFsProds.Col = 4
    flexNFsProds.Text = "Cliente"
    flexNFsProds.Col = 5
    flexNFsProds.Text = "Cidade"
    flexNFsProds.Col = 6
    flexNFsProds.Text = "UF"
    flexNFsProds.Col = 7
    flexNFsProds.Text = "Cod.Material"
    flexNFsProds.Col = 8
    flexNFsProds.Text = "Material"
    flexNFsProds.Col = 9
    flexNFsProds.Text = "Qtde."
    flexNFsProds.Col = 10
    flexNFsProds.Text = "Emissão NF"
    
    flexNFsProds.ColWidth(0) = 200
    flexNFsProds.ColWidth(1) = 1000
    flexNFsProds.ColWidth(2) = 1000
    flexNFsProds.ColWidth(3) = 1000
    flexNFsProds.ColWidth(4) = 2500
    flexNFsProds.ColWidth(5) = 2000
    flexNFsProds.ColWidth(6) = 300
    flexNFsProds.ColWidth(7) = 1000
    flexNFsProds.ColWidth(8) = 4000
    flexNFsProds.ColWidth(9) = 600
    flexNFsProds.ColWidth(10) = 1000
    
    flexNFsProds.ColAlignment(1) = 1
    flexNFsProds.ColAlignment(2) = 1
    flexNFsProds.ColAlignment(3) = 1
    flexNFsProds.ColAlignment(4) = 1
    flexNFsProds.ColAlignment(5) = 1
    flexNFsProds.ColAlignment(6) = 1
    flexNFsProds.ColAlignment(7) = 1
    flexNFsProds.ColAlignment(8) = 1
    flexNFsProds.ColAlignment(9) = 1
    flexNFsProds.ColAlignment(10) = 1
    
    If de_informa.rsSel_Estudios.State = 1 Then de_informa.rsSel_Estudios.Close
    de_informa.Sel_Estudios
    
    cmbEstudios.Clear
    cmbEstudios.AddItem "Todos"
    
    Do Until de_informa.rsSel_Estudios.EOF
    
        cmbEstudios.AddItem de_informa.rsSel_Estudios.Fields("id_estudio") & "-" & de_informa.rsSel_Estudios.Fields("estudio")
        de_informa.rsSel_Estudios.MoveNext
    
    Loop
    
    cmbEstudios.Text = "Todos"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
    
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
        If CDate(mskPer1.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
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
        If CDate(mskPer2.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub optCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optEmitidos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optCliCnpj_Click()
    If optCliTodos.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        txtFiltroCliente = ""
    ElseIf optCliCnpj.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Núm. CNPJ:"
    ElseIf optCliCodSap.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Código SAP:"
    ElseIf optCliFantasia.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Nome Fantasia:"
    ElseIf optCliTipoGrupo.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        txtFiltroCliente = ""
    End If

End Sub

Private Sub optCliCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optCliCodSap_Click()
     If optCliTodos.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        txtFiltroCliente = ""
    ElseIf optCliCnpj.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Núm. CNPJ:"
    ElseIf optCliCodSap.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Código SAP:"
    ElseIf optCliFantasia.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Nome Fantasia:"
    ElseIf optCliTipoGrupo.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        txtFiltroCliente = ""
    End If

End Sub

Private Sub optCliCodSap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optCliFantasia_Click()
    If optCliTodos.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        txtFiltroCliente = ""
    ElseIf optCliCnpj.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Núm. CNPJ:"
    ElseIf optCliCodSap.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Código SAP:"
    ElseIf optCliFantasia.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Nome Fantasia:"
    ElseIf optCliTipoGrupo.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        txtFiltroCliente = ""
    End If

End Sub

Private Sub optCliFantasia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optCliTipoGrupo_Click()
    If optCliTodos.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        txtFiltroCliente = ""
    ElseIf optCliCnpj.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Núm. CNPJ:"
    ElseIf optCliCodSap.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Código SAP:"
    ElseIf optCliFantasia.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Nome Fantasia:"
    ElseIf optCliTipoGrupo.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        txtFiltroCliente = ""
    End If

End Sub

Private Sub optCliTipoGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optCliTodos_Click()
    If optCliTodos.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = False
        txtFiltroCliente = ""
    ElseIf optCliCnpj.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Núm. CNPJ:"
    ElseIf optCliCodSap.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Código SAP:"
    ElseIf optCliFantasia.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        txtFiltroCliente = ""
        lblfiltrocliente.Caption = "Nome Fantasia:"
    ElseIf optCliTipoGrupo.Value = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        txtFiltroCliente = ""
    End If

End Sub

Private Sub optCliTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optEstudio_Click()
    If optEstudio.Value = True Then
        txtDestCGC.Enabled = True
        txtDestCGC.BackColor = xamarelo1
        cmdBuscaDest.Enabled = True
        txtMaterial.Enabled = True
        txtMaterial.BackColor = xamarelo1
        txtPacote.Enabled = True
        txtPacote.BackColor = xamarelo1
        tabVideolar.TabEnabled(1) = True
     ElseIf optVideolar.Value = True Then
        txtDestCGC.Enabled = False
        txtDestCGC.BackColor = xbranco
        cmdBuscaDest.Enabled = False
        txtMaterial.Enabled = False
        txtMaterial.BackColor = xbranco
        txtPacote.Enabled = False
        txtPacote.BackColor = xbranco
        tabVideolar.Tab = 0
    End If
End Sub

Private Sub optEstudio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Option4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optVideolar_Click()
    If optEstudio.Value = True Then
        txtDestCGC.Enabled = True
        txtDestCGC.BackColor = xamarelo1
        cmdBuscaDest.Enabled = True
        txtMaterial.Enabled = True
        txtMaterial.BackColor = xamarelo1
        txtPacote.Enabled = True
        txtPacote.BackColor = xamarelo1
        tabVideolar.TabEnabled(1) = True
     ElseIf optVideolar.Value = True Then
        txtDestCGC.Enabled = False
        txtDestCGC.BackColor = xbranco
        cmdBuscaDest.Enabled = False
        txtMaterial.Enabled = False
        txtMaterial.BackColor = xbranco
        txtPacote.Enabled = False
        txtPacote.BackColor = xbranco
        tabVideolar.Tab = 0
    End If
End Sub

Private Sub optVideolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    flexVideolar.ToolTipText = ""
End Sub

Private Sub txtRemetCGC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtRemetCGC_LostFocus()
    If Len(Trim$(txtRemetCGC)) > 5 Then
        If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
        de_informa.Sel_CadCliCGC Trim$(txtRemetCGC)
        
        If de_informa.rsSel_CadCliCGC.RecordCount < 1 Then
            lblRemetNome = ""
            MsgBox "CNPJ Não Encontrado !", vbExclamation, "Erro"
            Exit Sub
        Else
            lblRemetNome = de_informa.rsSel_CadCliCGC.Fields("nome")
        End If
    Else
        lblRemetNome = ""
    End If
    
End Sub

Private Sub txtDestCGC_Change()
    If Not IsNumeric(txtDestCGC) Then
        SendKeys "{BS}"
    End If
End Sub

Private Sub txtDestCGC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtDestCGC_LostFocus()
    If Len(Trim$(txtDestCGC)) > 0 Then
        If de_informa.rsSel_VLBuscaCliCodigo.State = 1 Then de_informa.rsSel_VLBuscaCliCodigo.Close
        de_informa.Sel_VLBuscaCliCodigo zeros(CDbl(txtDestCGC), 9)
        
        If de_informa.rsSel_VLBuscaCliCodigo.RecordCount > 0 Then
            lblDestNome.Caption = de_informa.rsSel_VLBuscaCliCodigo.Fields("clientenf")
        Else
            lblDestNome.Caption = "NÃO ENCONTRADO !"
            txtDestCGC.SetFocus
        End If
    Else
            lblDestNome.Caption = ""
    End If
End Sub

Private Sub txtMaterial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtMaterial_LostFocus()
    txtMaterial.Text = UCase(txtMaterial)
End Sub

Private Sub txtPacote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPacote_LostFocus()
    txtPacote.Text = UCase(txtPacote)
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
