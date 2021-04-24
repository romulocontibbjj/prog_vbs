VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsOcorr 
   Caption         =   "Consulta Ocorrências"
   ClientHeight    =   7635
   ClientLeft      =   1950
   ClientTop       =   660
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   8130
   Begin VB.CommandButton cmdImprTela 
      Height          =   375
      Left            =   6480
      Picture         =   "frmConsOcorr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton cmdRastrPrazo 
      Caption         =   "Rastrear Prazo Entrega..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdScanner 
      Caption         =   "POD CTC Scanner..."
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   7200
      TabIndex        =   24
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton cmdLancaPod 
      Caption         =   "Nova Ocorrência ou Baixa..."
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   7200
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid GridConsOcorr 
      Bindings        =   "frmConsOcorr.frx":0772
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      DataMember      =   "Sel_ConsOcorr2"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "data"
         Caption         =   "data"
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
         DataField       =   "hora"
         Caption         =   "hora"
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
         DataField       =   "cod_ocorr"
         Caption         =   "cd"
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
         DataField       =   "descr_ocorr"
         Caption         =   "ocorrência / descrição"
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
         DataField       =   "usu_ocorr"
         Caption         =   "usuário"
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
      BeginProperty Column05 
         DataField       =   "usu_dataocorr"
         Caption         =   "data inclusão"
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
      BeginProperty Column06 
         DataField       =   "obs_ocorr"
         Caption         =   "obs_ocorr"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   480,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   269,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3135,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1530,142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   30,047
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame12 
      Caption         =   "Informações sobre Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   7935
      Begin VB.Frame Frame2 
         Caption         =   "Env.Setor Arquivo"
         Height          =   975
         Left            =   6120
         TabIndex        =   38
         Top             =   1320
         Width           =   1695
         Begin VB.Label lblNumProt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   42
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Num."
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblArquivo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraPrazos 
         Caption         =   "Prazo - dias úteis"
         Height          =   1035
         Left            =   6120
         TabIndex        =   26
         Top             =   240
         Width           =   1695
         Begin VB.Label lblMetaPrazo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblTabPrazo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   600
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Meta/Prazo:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tab:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Pré - Baixa"
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2955
         Begin VB.Label lblDtUsuPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label lblDtBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label lblHsBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   16
            Top             =   600
            Width           =   750
         End
         Begin VB.Label lblRecebPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   17
            Top             =   960
            Width           =   2130
         End
         Begin VB.Label lblUsu_bxpre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quem Baixou:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   390
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   525
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Baixa Física"
         Height          =   2055
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   2955
         Begin VB.Label lblDtUsu 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   36
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label lblUsu_bx 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   21
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblReceb 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   20
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label lblHsBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   19
            Top             =   600
            Width           =   750
         End
         Begin VB.Label lblDtBaixa 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   18
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Quem Baixou:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   630
            Width           =   390
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   525
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Observação de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblObsEntr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   7680
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observação de Ocorrência"
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
      TabIndex        =   0
      Top             =   1800
      Width           =   7935
      Begin VB.Label lblObs_Ocorr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   960
         Left            =   105
         TabIndex        =   14
         Top             =   315
         Width           =   7680
      End
   End
End
Attribute VB_Name = "frmConsOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdLancaPod_Click()
    If Mid$(xdireitos, 13, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        'desvio para o form de POD
        Me.Hide
        frmPod.Show
        frmPod.txtfilial.Text = frmSac.txtfilial.Text
        frmPod.txtCTC.Text = frmSac.txtCTC.Text
        DoEvents
        Unload frmSac
        Unload Me
        DoEvents
        frmPod.cmdProcurar.SetFocus
        DoEvents
        SendKeys "{ENTER}"   'pressiona o Enter
    End If
End Sub

Private Sub cmdRastrPrazo_Click()
    de_informa.Alt_AtualPrazoSCTC transctc(frmSac.txtfilial, frmSac.txtCTC)
'    frmAtualPrazos.Show 1
    If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
    de_informa.Sel_ConsOcorr transctc(frmSac.txtfilial, frmSac.txtCTC), "01"
    frmRastrPrazo.lblFilialCTC = transctc(frmSac.txtfilial, frmSac.txtCTC)
    frmRastrPrazo.lblModal = frmSac.lblModal
    frmRastrPrazo.lblCidadeDest = frmSac.lblCidade_Dest
    frmRastrPrazo.lblUFdest = frmSac.lblUf_Dest
    frmRastrPrazo.lblEmissao = frmSac.lblData
    frmRastrPrazo.lblHsEmiss = frmSac.lblHora
    frmRastrPrazo.lblEntrega = de_informa.rsSel_ConsOcorr.Fields("data")
    frmRastrPrazo.lblHsEntr = de_informa.rsSel_ConsOcorr.Fields("hora")
    frmRastrPrazo.lblMeta = de_informa.rsSel_ConsOcorr.Fields("prazoentr")
    frmRastrPrazo.lblPrazo = de_informa.rsSel_ConsOcorr.Fields("diasuteis")
    frmConsOcorr.lblMetaPrazo = Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("prazoentr"))) & "/" & Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("diasuteis")))
    DoEvents
    frmRastrPrazo.Show 1
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdScanner_Click()
    Dim xarquivo As String
    xarquivo = frmSac.txtfilial.Text & frmSac.txtCTC.Text & ".JPG"
    If Dir(App.Path & "\IMAGENS\" & xarquivo) = "" Then
        MsgBox "Arquivo não Encontrado !", vbCritical, "Erro"
        Exit Sub
    Else
        frmImagem.Image1.Picture = LoadPicture(App.Path & "\IMAGENS\" & xarquivo)
        frmImagem.lblFilial.Caption = frmSac.txtfilial.Text
        frmImagem.lblCtc.Caption = frmSac.txtCTC.Text
        frmImagem.Show 1
    End If
End Sub

Private Sub Form_Activate()
    Me.Left = 1770
    Me.Top = 585
End Sub

Private Sub Form_Load()
    'fecha os recordsets caso estejam abertos
        If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
    'consulta que traz os campos = 01 que é dado de entrega (ENTREGA REALIZADA)
        de_informa.Sel_ConsOcorr transctc(frmSac.txtfilial.Text, frmSac.txtCTC.Text), "01"
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
    'atualiza os campos referente a dados de entrega
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")) = False Then lblDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")) = False Then lblHsBaixaPre = de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("recebpre")) = False Then lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")) = False Then lblUsu_bxpre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) = False Then lblDtBaixa = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("hsbaixa")) = False Then lblHsBaixa = de_informa.rsSel_ConsOcorr.Fields("hsbaixa")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("receb")) = False Then lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_bx")) = False Then lblUsu_bx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) = False Then lblObsEntr = de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_datapre")) = False Then lblDtUsuPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_databx")) = False Then lblDtUsu = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("rel_arq_data")) = False And de_informa.rsSel_ConsOcorr.Fields("rel_arquivo") = "S" Then
                lblArquivo = de_informa.rsSel_ConsOcorr.Fields("rel_arq_data")
                If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("rel_arq_num")) Then
                    lblNumProt = String(6 - Len(Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("rel_arq_num")))), "0") & Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("rel_arq_num")))
                End If
            End If
            lblTabPrazo = de_informa.rsSel_ConsCadCli.Fields("prazo")
            lblMetaPrazo = Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("prazoentr"))) & "/" & Trim$(Str(de_informa.rsSel_ConsOcorr.Fields("diasuteis")))
        End If
    'consulta que traz os campo que são dados de ocorrência
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
        de_informa.Sel_ConsOcorr2 transctc(frmSac.txtfilial.Text, frmSac.txtCTC.Text), "01"
        Set GridConsOcorr.DataSource = de_informa
        GridConsOcorr.DataMember = "Sel_ConsOcorr2"
        GridConsOcorr.Refresh
        If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr") <> "" Then
                lblObs_Ocorr = de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr")
            End If
        End If


End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmConsOcorr = Nothing
End Sub

Private Sub GridConsOcorr_Click()
    If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
    'atualiza o campo de obs de ocorrência quando clicado no grid
        lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
    End If
End Sub

