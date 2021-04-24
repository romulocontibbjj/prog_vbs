VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAIRRel 
   Caption         =   "Informa - Acompanhamento de Emissão de AWBs"
   ClientHeight    =   8295
   ClientLeft      =   -1575
   ClientTop       =   2355
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmAIRRel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.Frame FraResultado 
      Caption         =   "** Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   11775
      Begin VB.Frame FraMsg 
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   3720
         TabIndex        =   26
         Top             =   1860
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton CmdFecharMsg 
            Caption         =   "Ok"
            Height          =   315
            Left            =   1043
            TabIndex        =   30
            Top             =   2460
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label LblMsg 
            Alignment       =   2  'Center
            Caption         =   $"frmAIRRel.frx":000C
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2055
            Left            =   300
            TabIndex        =   29
            Top             =   300
            Visible         =   0   'False
            Width           =   4140
         End
         Begin VB.Label LblGerando 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gerando Arquivo... Aguarde..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   795
            Visible         =   0   'False
            Width           =   4230
         End
         Begin VB.Label LblGerando 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gerando Arquivo... Aguarde..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   270
            TabIndex        =   28
            Top             =   840
            Visible         =   0   'False
            Width           =   4230
         End
      End
      Begin VB.Frame FraMsg2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   3840
         TabIndex        =   31
         Top             =   1980
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CommandButton cmdGerarArq 
         Caption         =   "Geração de Arquivo de Movimento"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8580
         TabIndex        =   13
         Top             =   240
         Width           =   3015
      End
      Begin MSFlexGridLib.MSFlexGrid FlexAWB 
         Height          =   5715
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10081
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Label LblREGRET 
         AutoSize        =   -1  'True
         Caption         =   "Registros Retornados: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   330
         Width           =   2070
      End
   End
   Begin VB.Frame fraComandos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   5760
      TabIndex        =   17
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox ComboCIA 
         Height          =   315
         Left            =   3300
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox ComboREG 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3300
         TabIndex        =   14
         Top             =   900
         Width           =   2655
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processar / Atualizar Dados"
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   900
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar pela Cia. Aérea"
         Height          =   195
         Left            =   3300
         TabIndex        =   25
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar pela Região"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   1275
      End
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
      Height          =   1395
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5595
      Begin VB.Frame fraPorPeriodo 
         Caption         =   "Período"
         Height          =   1035
         Left            =   1620
         TabIndex        =   19
         Top             =   180
         Width           =   3855
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   2100
            TabIndex        =   8
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   420
            TabIndex        =   7
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "( Intervalo Máximo de 60 dias )"
            Height          =   195
            Left            =   825
            TabIndex        =   21
            Top             =   240
            Width           =   2160
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1860
            TabIndex        =   20
            Top             =   540
            Width           =   90
         End
      End
      Begin VB.Frame fraPorEmissao 
         Height          =   1035
         Left            =   1620
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   3855
         Begin VB.OptionButton opt60d 
            Caption         =   "Últimos 60 dias"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton opt15d 
            Caption         =   "Últimos 15 dias"
            Height          =   195
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton opt30d 
            Caption         =   "Últimos 30 dias"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.OptionButton optPorMes 
         Caption         =   "Por Mês..."
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optPorPeriodo 
         Caption         =   "Por Período..."
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optPorEmissao 
         Caption         =   "Emissão nos..."
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame fraPorMesAno 
         Height          =   1035
         Left            =   1620
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   3855
         Begin VB.ComboBox comboMesAnoAcomp 
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Text            =   "Mes/Ano"
            Top             =   420
            Width           =   3375
         End
      End
   End
End
Attribute VB_Name = "frmAIRRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdFecharMsg_Click()
FraMsg.Visible = False
FraMsg2.Visible = False
LblGerando(0).Visible = False
LblGerando(1).Visible = False
LblMsg.Visible = False
CmdFecharMsg.Visible = False
DoEvents
End Sub

Private Sub cmdGerarArq_Click()
FraMsg.Visible = True
FraMsg2.Visible = True
LblGerando(0).Visible = True
LblGerando(1).Visible = True
DoEvents

Dim xInf As Boolean
xarquivo = "c:\Resumo AWBS Emitidos - " & Year(Date) & Month(Date) & Day(Date) & " - " & Hour(Time) & Minute(Time) & Second(Time) & ".TXT"
xInf = False

Open xarquivo For Output As #1

    For Y = 0 To FlexAWB.Rows - 1
    xlinha = ""
        For X = 0 To FlexAWB.Cols - 1
        xlinha = xlinha & FlexAWB.TextMatrix(Y, X) & "#"
        Next
    Print #1, xlinha
    Next
Close #1

LblGerando(0).Visible = False
LblGerando(1).Visible = False
LblMsg.Visible = True
CmdFecharMsg.Visible = True
DoEvents

End Sub

Private Sub cmdInserirVOO_Click()
xACOMP = True
frmAcompAWB.Show 1
End Sub

Private Sub cmdProcessa_Click()
cmdProcessa.Enabled = False
cmdGerarArq.Enabled = False
cmdSair.Enabled = False
Me.MousePointer = 11
DoEvents

Dim DataInicial As String
Dim DataFinal As String
Dim xBarra As Integer
Dim Mes As Integer
Dim Ano As Integer
Dim UltimoDiaMes As Integer
Dim X As Integer
Dim Y As Integer


    If optPorPeriodo.Value = True Then
    DataInicial = mskPer1.Text
    DataFinal = mskPer2.Text
    ElseIf optPorMes.Value = True Then
    xBarra = InStr(1, comboMesAnoAcomp.Text, "/", vbTextCompare)
    Mes = NumMes(Mid(comboMesAnoAcomp.Text, 1, xBarra - 1))
    Ano = Val(Trim(Mid(comboMesAnoAcomp.Text, xBarra + 1)))
        If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
        UltimoDiaMes = 31
        ElseIf Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
        UltimoDiaMes = 30
        ElseIf Mes = 2 Then
            If (Ano / 4) = Int(Ano / 4) Then
            UltimoDiaMes = 29
            ElseIf (Ano / 4) <> Int(Ano / 4) Then
            UltimoDiaMes = 28
            End If
        End If
    DataInicial = "01/" & Mes & "/" & Ano
    DataFinal = UltimoDiaMes & "/" & Mes & "/" & Ano
    ElseIf optPorEmissao.Value = True Then
        If opt15d.Value = True Then
        DataInicial = Date - 15
        DataFinal = Date
        ElseIf opt30d.Value = True Then
        DataInicial = Date - 30
        DataFinal = Date
        ElseIf opt60d.Value = True Then
        DataInicial = Date - 60
        DataFinal = Date
        End If
    End If
    
    Dim Sigla As String
    
    xBarra = InStr(1, ComboREG.Text, "-", vbTextCompare)
    Sigla = Trim(UCase(Mid(ComboREG, xBarra + 1)))
    xCia = Mid(ComboCIA, 1, 2)

    If de_informa.rsSelAWBPeriodo.State = 1 Then de_informa.rsSelAWBPeriodo.Close
    de_informa.selawbperiodo CDate(DataInicial), CDate(DataFinal), xCia & "%", Sigla & "%"
    

    
    If de_informa.rsSelAWBPeriodo.RecordCount = 0 Then
    MsgBox "Sua pesquisa para este período não retornou registro algum.", vbCritical, ""
    cmdProcessa.Enabled = True
    cmdGerarArq.Enabled = True
    cmdSair.Enabled = True
    Me.MousePointer = 0
    Exit Sub
    End If

FlexAWB.Clear
    
LblREGRET.Caption = "Registros Retornados: " & de_informa.rsSelAWBPeriodo.RecordCount
FlexAWB.Rows = de_informa.rsSelAWBPeriodo.RecordCount + 1
FlexAWB.Cols = 8
FlexAWB.FixedRows = 1
FlexAWB.FixedCols = 0
   

FlexAWB.TextMatrix(0, 0) = "Data"
FlexAWB.TextMatrix(0, 1) = "Nome Cia."
FlexAWB.TextMatrix(0, 2) = "Cia."
FlexAWB.TextMatrix(0, 3) = "AWBs"
FlexAWB.TextMatrix(0, 4) = "Volumes"
FlexAWB.TextMatrix(0, 5) = "Peso Real"
FlexAWB.TextMatrix(0, 6) = "Peso Cubado"
FlexAWB.TextMatrix(0, 7) = "Frete Nacional"

FlexAWB.ColWidth(0) = 1300
FlexAWB.ColWidth(1) = 1800
FlexAWB.ColWidth(2) = 1300
FlexAWB.ColWidth(3) = 1300
FlexAWB.ColWidth(4) = 1300
FlexAWB.ColWidth(5) = 1300
FlexAWB.ColWidth(6) = 1300
FlexAWB.ColWidth(7) = 1300

X = 0

    Do Until de_informa.rsSelAWBPeriodo.EOF
    X = X + 1
    FlexAWB.TextMatrix(X, 0) = X
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("data")) = False Then FlexAWB.TextMatrix(X, 0) = de_informa.rsSelAWBPeriodo.Fields("data")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("nomecia")) = False Then FlexAWB.TextMatrix(X, 1) = de_informa.rsSelAWBPeriodo.Fields("nomecia")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("cia")) = False Then FlexAWB.TextMatrix(X, 2) = de_informa.rsSelAWBPeriodo.Fields("cia")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("AWBs")) = False Then FlexAWB.TextMatrix(X, 3) = de_informa.rsSelAWBPeriodo.Fields("awbs")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("volumes")) = False Then FlexAWB.TextMatrix(X, 4) = de_informa.rsSelAWBPeriodo.Fields("volumes")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("peso_real")) = False Then FlexAWB.TextMatrix(X, 5) = de_informa.rsSelAWBPeriodo.Fields("peso_real")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("peso_cubado")) = False Then FlexAWB.TextMatrix(X, 6) = de_informa.rsSelAWBPeriodo.Fields("peso_cubado")
    If IsNull(de_informa.rsSelAWBPeriodo.Fields("frete_nacional")) = False Then FlexAWB.TextMatrix(X, 7) = Format(de_informa.rsSelAWBPeriodo.Fields("frete_nacional"), "##,##0.00")
    DoEvents
    de_informa.rsSelAWBPeriodo.MoveNext
    Loop
    
cmdProcessa.Enabled = True
cmdGerarArq.Enabled = True
cmdSair.Enabled = True
Me.MousePointer = 0
DoEvents

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
optPorPeriodo_Click
mskPer1.SetFocus
End Sub

Private Sub Form_Load()
Dim MesAtual As Integer, AnoAtual As Integer, Retro As Integer, X As Integer

FlexAWB.Rows = 0
comboMesAnoAcomp.Clear
MesAtual = Month(Date)
AnoAtual = Year(Date)
Retro = 18

comboMesAnoAcomp.AddItem PriMaiuscula(NomeMes(MesAtual)) & "/" & AnoAtual
comboMesAnoAcomp.Text = PriMaiuscula(NomeMes(MesAtual)) & "/" & AnoAtual
    
    For X = 0 To Retro
    MesAtual = MesAtual - 1
        If MesAtual < 1 Then
        AnoAtual = AnoAtual - 1
        MesAtual = 12
        End If
    
    comboMesAnoAcomp.AddItem PriMaiuscula(NomeMes(MesAtual)) & "/" & AnoAtual
    Next

If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
de_informa.SelAeroportoCidade "%"

ComboREG.Clear
    Do Until de_informa.rsSelAeroportoCidade.EOF
    ComboREG.AddItem PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & UCase(Trim(de_informa.rsSelAeroportoCidade.Fields("sigla")))
    de_informa.rsSelAeroportoCidade.MoveNext
    Loop

If de_informa.rsSelCiaAerea.State = 1 Then de_informa.rsSelCiaAerea.Close
de_informa.SelCiaAerea "%"

ComboCIA.Clear
    Do Until de_informa.rsSelCiaAerea.EOF
    ComboCIA.AddItem UCase(de_informa.rsSelCiaAerea.Fields("codcia")) & " - " & PriMaiuscula(Trim(de_informa.rsSelCiaAerea.Fields("fantasia")))
    de_informa.rsSelCiaAerea.MoveNext
    Loop

End Sub


Private Sub mskPer1_GotFocus()
Call Date_MskEdBox_GotFocus(mskPer1)
End Sub

Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer1_LostFocus()
Call Date_MskEdBox_LostFocus(mskPer1)
End Sub

Private Sub mskper2_GotFocus()
Call Date_MskEdBox_GotFocus(mskPer2)
End Sub

Private Sub mskper2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskper2_LostFocus()
Call Date_MskEdBox_LostFocus(mskPer2)
End Sub

Private Sub opt30d_Click()
    mskPer1.Mask = ""
    mskPer1.Text = ""
    mskPer1.Mask = "##/##/####"
    mskPer2.Mask = ""
    mskPer2.Text = ""
    mskPer2.Mask = "##/##/####"
End Sub

Private Sub opt30d_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub


Private Sub optPorEmissao_Click()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If
End Sub

Private Sub optPorEmissao_GotFocus()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If
End Sub

Private Sub optPorMes_Click()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If

End Sub

Private Sub optPorPeriodo_Click()
    If optPorEmissao.Value = True Then
        fraPorEmissao.Visible = True
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = False
    ElseIf optPorPeriodo.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = True
        fraPorMesAno.Visible = False
    ElseIf optPorMes.Value = True Then
        fraPorEmissao.Visible = False
        fraPorPeriodo.Visible = False
        fraPorMesAno.Visible = True
    End If

End Sub

