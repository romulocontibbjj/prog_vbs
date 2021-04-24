VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAcompInformacao 
   Caption         =   "Acompanhamento da Informação"
   ClientHeight    =   7515
   ClientLeft      =   1845
   ClientTop       =   1485
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Prioridades"
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
      Left            =   7920
      TabIndex        =   18
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optPriori 
         Caption         =   "Somente Prioridades"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optUrg 
         Caption         =   "Somente Urgências"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optTodo 
         Caption         =   "Todo Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   495
      Left            =   10320
      Picture         =   "frmAcompInformacao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   11040
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdProcessar 
      Caption         =   "Processar ..."
      Height          =   375
      Left            =   10320
      TabIndex        =   14
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "** Período"
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
      TabIndex        =   2
      Top             =   120
      Width           =   7695
      Begin VB.Frame Frame6 
         Caption         =   "No Período de ...  (máximo de 30 dias)"
         Height          =   735
         Left            =   4560
         TabIndex        =   7
         Top             =   360
         Width           =   3015
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   1680
            TabIndex        =   10
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
            TabIndex        =   8
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1440
            TabIndex        =   9
            Top             =   360
            Width           =   90
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Emissão..."
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3735
         Begin VB.OptionButton opt60d 
            Caption         =   "Últ. 60 dias"
            Height          =   195
            Left            =   1560
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton opt30d 
            Caption         =   "Últ. 30 dias"
            Height          =   195
            Left            =   2040
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optPer15d 
            Caption         =   "Últ. 15 dias"
            Height          =   195
            Left            =   600
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "OU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
   End
   Begin TabDlg.SSTab tabInformacao 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Posição Geral"
      TabPicture(0)   =   "frmAcompInformacao.frx":0772
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridPosicao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Movimento Entregue"
      TabPicture(1)   =   "frmAcompInformacao.frx":078E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridPosicao 
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5530
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmAcompInformacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdProcessar_Click()
    Dim xdata1 As Date, xdata2 As Date, xcont As Integer, xufs As String, xpriori As String
    
    Me.MousePointer = 11
    
    'trata as datas
    
    If optPer15d.Value = True Then
        mskPer1.Mask = ""
        mskPer1.Text = ""
        mskPer1.Mask = "##/##/####"
        mskPer2.Mask = ""
        mskPer2.Text = ""
        mskPer2.Mask = "##/##/####"
        xdata1 = datahora("data") - 15
        xdata2 = datahora("data")
    ElseIf opt30d.Value = True Then
        mskPer1.Mask = ""
        mskPer1.Text = ""
        mskPer1.Mask = "##/##/####"
        mskPer2.Mask = ""
        mskPer2.Text = ""
        mskPer2.Mask = "##/##/####"
        xdata1 = datahora("data") - 30
        xdata2 = datahora("data")
    ElseIf opt60d.Value = True Then
        mskPer1.Mask = ""
        mskPer1.Text = ""
        mskPer1.Mask = "##/##/####"
        mskPer2.Mask = ""
        mskPer2.Text = ""
        mskPer2.Mask = "##/##/####"
        xdata1 = datahora("data") - 60
        xdata2 = datahora("data")
    Else
        If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
            MsgBox "Período Escolhido Inválido !"
            mskPer1.SetFocus
            Me.MousePointer = 0
            Exit Sub
        End If
        If CDate(mskPer1) > CDate(mskPer2) Then
            MsgBox "Período de Escolha Inválido ! Data Início Maior que a Data Final."
            mskPer1.SetFocus
            Me.MousePointer = 0
            Exit Sub
        End If
        xdata1 = CDate(mskPer1)
        xdata2 = CDate(mskPer2)
    End If
    
    If xdata2 - xdata1 > 32 Then
        MsgBox "Período Escolhido Maior que 30 Dias ! Escolha um Período Menor."
        mskPer1.SetFocus
        Me.MousePointer = 0
        Exit Sub
    End If
    
    gridPosicao.Clear
    
    gridPosicao.Row = 0
    
    gridPosicao.Col = 1
    gridPosicao.Text = "Cod.Reg."
    gridPosicao.Col = 2
    gridPosicao.Text = "Responsável"
    gridPosicao.Col = 3
    gridPosicao.Text = "Abrangência UFs"
    gridPosicao.Col = 4
    gridPosicao.Text = "CTCs"
    gridPosicao.Col = 5
    gridPosicao.Text = "Sem.Pos."
    gridPosicao.Col = 6
    gridPosicao.Text = "Em Ocorr."
    gridPosicao.Col = 7
    gridPosicao.Text = "Trânsito"
    gridPosicao.Col = 8
    gridPosicao.Text = "Entregues"
    gridPosicao.Col = 9
    gridPosicao.Text = "% Pend."
    
    'PRIORIDADES
    
    If optTodo = True Then
        xpriori = "%"
    ElseIf optUrg = True Then
        xpriori = "URGÊNCIA%"
    ElseIf optPriori = True Then
        xpriori = "PRIORIDADE%"
    End If
    
    'buscas as regiões
    
    If de_informa.rsSel_AcompInfRegioes.State = 1 Then de_informa.rsSel_AcompInfRegioes.Close
    de_informa.Sel_AcompInfRegioes
    
    gridPosicao.Rows = de_informa.rsSel_AcompInfRegioes.RecordCount + 1
    
    'monta a grid por região
    
    For xcont = 1 To de_informa.rsSel_AcompInfRegioes.RecordCount
    
        gridPosicao.Row = xcont
        
        'REGIAO
        
        gridPosicao.Col = 1
        gridPosicao.Text = de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        
        'RESPONSÁVEL
        
        gridPosicao.Col = 2
        gridPosicao.Text = de_informa.rsSel_AcompInfRegioes.Fields("atendsac")
        
        'ABRANGÊNCIA
        
        gridPosicao.Col = 3
        If de_informa.rsSel_AcompInfUFs.State = 1 Then de_informa.rsSel_AcompInfUFs.Close
        de_informa.Sel_AcompInfUFs de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        xufs = ""
        Do Until de_informa.rsSel_AcompInfUFs.EOF
            xufs = xufs & de_informa.rsSel_AcompInfUFs.Fields("uf") & ","
            de_informa.rsSel_AcompInfUFs.MoveNext
        Loop
        
        gridPosicao.Text = xufs
        
        'QTDE DE CTCs
        
        If de_informa.rsSel_AcompInfCtcs.State = 1 Then de_informa.rsSel_AcompInfCtcs.Close
        de_informa.Sel_AcompInfCtcs xdata1, xdata2, de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xpriori
        
        gridPosicao.Col = 4
        gridPosicao.Text = de_informa.rsSel_AcompInfCtcs.Fields("qtd")
    
        'SEM POSIÇÃO
        
        If de_informa.rsSel_AcompInfSemPos.State = 1 Then de_informa.rsSel_AcompInfSemPos.Close
        de_informa.Sel_AcompInfSemPos xdata1, xdata2, de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xpriori
        
        gridPosicao.Col = 5
        gridPosicao.Text = de_informa.rsSel_AcompInfSemPos.Fields("qtd")
        
        'EM OCORRÊNCIA
    
        If de_informa.rsSel_AcompInfEmOcorr.State = 1 Then de_informa.rsSel_AcompInfEmOcorr.Close
        de_informa.Sel_AcompInfEmOcorr xdata1, xdata2, de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xpriori
        
        gridPosicao.Col = 6
        gridPosicao.Text = de_informa.rsSel_AcompInfEmOcorr.Fields("qtd")
        
        'EM TRÂNSITO
        
        If de_informa.rsSel_AcompInfTransito.State = 1 Then de_informa.rsSel_AcompInfTransito.Close
        de_informa.Sel_AcompInfTransito xdata1, xdata2, de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xpriori
    
        gridPosicao.Col = 7
        gridPosicao.Text = de_informa.rsSel_AcompInfTransito.Fields("qtd")
    
        'ENTREGUES
    
        If de_informa.rsSel_AcompInfEntregues.State = 1 Then de_informa.rsSel_AcompInfEntregues.Close
        de_informa.Sel_AcompInfEntregues xdata1, xdata2, de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xpriori
    
        gridPosicao.Col = 8
        gridPosicao.Text = de_informa.rsSel_AcompInfEntregues.Fields("qtd")
        
        'PERCENTUAL PENDÊNCIAS
    
        gridPosicao.Col = 9
        If de_informa.rsSel_AcompInfCtcs.Fields("qtd") > 0 Then
            gridPosicao.Text = Format((de_informa.rsSel_AcompInfSemPos.Fields("qtd") + _
                            de_informa.rsSel_AcompInfEmOcorr.Fields("qtd")) / _
                            de_informa.rsSel_AcompInfCtcs.Fields("qtd"), "##0.0%")
        Else
            gridPosicao.Text = "0"
        End If
        
        DoEvents
        
        de_informa.rsSel_AcompInfRegioes.MoveNext
    
    Next
        
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    gridPosicao.Cols = 10
    gridPosicao.ColWidth(0) = 200
    gridPosicao.ColWidth(1) = 750
    gridPosicao.ColWidth(2) = 1100
    gridPosicao.ColWidth(3) = 3550
    gridPosicao.ColWidth(4) = 900
    gridPosicao.ColWidth(5) = 900
    gridPosicao.ColWidth(6) = 900
    gridPosicao.ColWidth(7) = 900
    gridPosicao.ColWidth(8) = 900
    gridPosicao.ColWidth(9) = 900
    
    gridPosicao.Row = 0
    
    gridPosicao.Col = 1
    gridPosicao.Text = "Cod.Reg."
    gridPosicao.Col = 2
    gridPosicao.Text = "Responsável"
    gridPosicao.Col = 3
    gridPosicao.Text = "Abrangência UFs"
    gridPosicao.Col = 4
    gridPosicao.Text = "CTCs"
    gridPosicao.Col = 5
    gridPosicao.Text = "Sem.Pos."
    gridPosicao.Col = 6
    gridPosicao.Text = "Em Ocorr."
    gridPosicao.Col = 7
    gridPosicao.Text = "Trânsito"
    gridPosicao.Col = 8
    gridPosicao.Text = "Entregues"
    gridPosicao.Col = 9
    gridPosicao.Text = "% Pend."
    
    
    
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAcompInformacao = Nothing
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
        If IsDate(mskPer2.Text) Then
'            If CDate(mskPer2.Text) < CDate(mskPer1.Text) Then
'                MsgBox "Período Inválido !", vbCritical, "Erro"
'                mskPer1.SetFocus
'                Exit Sub
'            Else
                opt30d.Value = False
                opt60d.Value = False
                optPer15d.Value = False
'            End If
        End If
    Else
        If mskPer2.Text = "__/__/____" And opt30d = False And opt60d = False And optPer15d = False Then
            optPer15d.Value = True
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
        If IsDate(mskPer1.Text) Then
'            If CDate(mskPer2.Text) < CDate(mskPer1.Text) Then
'                MsgBox "Período Inválido !", vbCritical, "Erro"
'                'mskPer2.SetFocus
'                Exit Sub
'            Else
                opt30d.Value = False
                opt60d.Value = False
                optPer15d.Value = False
'            End If
        End If
    Else
        If mskPer1.Text = "__/__/____" And opt30d = False And opt60d = False And optPer15d = False Then
            optPer15d.Value = True
        End If
    End If
End Sub

Private Sub optPriori_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optTodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optUrg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
