VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Environment Organizer"
   ClientHeight    =   7605
   ClientLeft      =   2535
   ClientTop       =   1620
   ClientWidth     =   10440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   435
      Left            =   8520
      TabIndex        =   12
      Top             =   6780
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CDWin 
      Left            =   9300
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdGerar 
      Caption         =   "Gerar Novo Arquivo"
      Height          =   435
      Left            =   8520
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Commando"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   "Campo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   6795
      Left            =   7980
      TabIndex        =   2
      Top             =   420
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   11986
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "Ordenar Comandos"
      Height          =   435
      Left            =   8520
      TabIndex        =   1
      Top             =   5220
      Width           =   1815
   End
   Begin VB.CommandButton CmdCarregar 
      Caption         =   "Carregar Comandos"
      Height          =   435
      Left            =   8520
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tree dos Comandos"
      TabPicture(0)   =   "frmMain.frx":0CE6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Estrutura .DSR"
      TabPicture(1)   =   "frmMain.frx":0D02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlexLinha"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Nome Comandos"
      TabPicture(2)   =   "frmMain.frx":0D1E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FlexOrdena"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   ".DSR Ordenado"
         Height          =   6435
         Left            =   3900
         TabIndex        =   9
         Top             =   540
         Width           =   3735
         Begin MSComctlLib.TreeView TVOrdenada 
            Height          =   6075
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   10716
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   ".DSR Original"
         Height          =   6435
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   3735
         Begin MSComctlLib.TreeView TVOriginal 
            Height          =   6075
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   10716
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FlexLinha 
         Height          =   6555
         Left            =   -74880
         TabIndex        =   7
         Top             =   420
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   11562
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid FlexOrdena 
         Height          =   6555
         Left            =   -74880
         TabIndex        =   8
         Top             =   420
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   11562
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
      End
   End
   Begin VB.Label LblArquivoOriginal 
      Height          =   1875
      Left            =   8520
      TabIndex        =   13
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label LblMensagem 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "                                                                                                      "
      Height          =   195
      Left            =   1695
      TabIndex        =   11
      Top             =   7320
      Width           =   4605
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCarregar_Click()
Dim Linha As String, xKey As String
Dim Nivel As Integer
Dim DentroCon As Boolean, DentroRS As Boolean, DentroCampo As Boolean, DentroPar As Boolean
Dim NomeComando As String, NomeCampo As String
Dim NumCon As Integer, NumRS As Integer, NumFields As Integer, NumPar As Integer
Dim InicioCampo As Long, FimCampo As Long, InicioRS As Long, FimRS As Long, InicioPar As Long, FimPar As Long, InicioCon As Long, FimCon As Long
Dim NoRS As Long
Dim ContLinha As Long
Dim IndexRS As Long, IndexCampo As Long, IndexCon As Long, IndexPar As Long

Dim xArquivoOriginal As String

CDWin.DialogTitle = "Escolha o Arquivo que deseja organizar"
CDWin.ShowOpen

xArquivoOriginal = CDWin.FileName

If xArquivoOriginal = "" Then Exit Sub
    If UCase(Mid(xArquivoOriginal, Len(xArquivoOriginal) - 2)) <> "DSR" Then
    MsgBox "Tipo Inválido de Arquivo!", vbCritical, ""
    Exit Sub
    End If

LblArquivoOriginal.Caption = xArquivoOriginal

CmdCarregar.Enabled = False
CmdOrdenar.Enabled = False
CmdGerar.Enabled = False
CmdSair.Enabled = False
Me.MousePointer = 11
LblMensagem.Caption = "Aguarde. Lendo o arquivo *.DSR."
DoEvents

TVOriginal.Nodes.Clear
TVOrdenada.Nodes.Clear
FlexLinha.Clear
FlexOrdena.Clear
FlexLinha.Rows = 0
'FlexLinha.Cols = 0
FlexOrdena.Rows = 0

Open xArquivoOriginal For Input As #1

Nivel = 1
ContLinha = 0

    Do Until EOF(1)
    Line Input #1, Linha
    ContLinha = ContLinha + 1
    'Linha = Trim(Linha)
    
        If InStr(1, Linha, "BeginProperty", vbTextCompare) > 0 Then
        
            If InStr(1, Linha, "BeginProperty Connection", vbTextCompare) Then
            DentroCon = True
            IndexCon = CDbl(Trim(Mid(Trim(Linha), 25)))
            InicioCon = ContLinha
            ElseIf InStr(1, Linha, "BeginProperty Recordset", vbTextCompare) Then
            DentroRS = True
            IndexRS = CDbl(Trim(Mid(Trim(Linha), 24)))
            InicioRS = ContLinha
            ElseIf InStr(1, Linha, "BeginProperty Field", vbTextCompare) Then
            DentroCampo = True
            IndexCampo = CDbl(Trim(Mid(Trim(Linha), 20)))
            InicioCampo = ContLinha
            ElseIf InStr(1, Linha, "BeginProperty P", vbTextCompare) Then
            DentroPar = True
            IndexPar = CDbl(Trim(Mid(Trim(Linha), 16)))
            InicioPar = ContLinha
            End If
        Nivel = Nivel + 1
        FlexLinha.Rows = FlexLinha.Rows + 1
        
            If FlexLinha.Cols < Nivel Then
            FlexLinha.Cols = Nivel
            End If
            
        FlexLinha.TextMatrix(FlexLinha.Rows - 1, Nivel - 1) = Linha
        
        ElseIf InStr(1, Linha, "EndProperty", vbTextCompare) > 0 Then
                    
            If DentroPar = True Then
            DentroPar = False
            FimPar = ContLinha
            ElseIf DentroCampo = True Then
            DentroCampo = False
            FimCampo = ContLinha
            ElseIf DentroRS = True Then
            DentroRS = False
            FimRS = ContLinha
            TVOriginal.Nodes.Item(NoRS).Key = TVOriginal.Nodes.Item(NoRS).Key & String(6 - Len(Trim(Str(FimRS))), "0") & Trim(Str(FimRS))
            ElseIf DentroCon = True Then
            DentroCon = False
            FimCon = ContLinha
            End If
        
        
        FlexLinha.Rows = FlexLinha.Rows + 1
        FlexLinha.TextMatrix(FlexLinha.Rows - 1, Nivel - 1) = Linha
        Nivel = Nivel - 1
        Else
        
            If InStr(1, Linha, "NumConnections", vbTextCompare) > 0 And DentroCon = False And DentroRS = False And DentroCampo = False And DentroPar = False Then
            NumCon = Val(Trim(Mid(Linha, InStr(1, Linha, "=", vbTextCompare) + 1)))
            ElseIf InStr(1, Linha, "NumRecordsets", vbTextCompare) > 0 And DentroCon = False And DentroRS = False And DentroCampo = False And DentroPar = False Then
            NumRS = Val(Trim(Mid(Linha, InStr(1, Linha, "=", vbTextCompare) + 1)))
            ElseIf InStr(1, Linha, "NumFields", vbTextCompare) > 0 And DentroCon = False And DentroRS = True And DentroCampo = False And DentroPar = False Then
            NumFields = Val(Trim(Mid(Linha, InStr(1, Linha, "=", vbTextCompare) + 1)))
            ElseIf InStr(1, Linha, "ParamCount", vbTextCompare) > 0 And DentroCon = False And DentroRS = True And DentroCampo = False And DentroPar = False Then
            NumPar = Val(Trim(Mid(Linha, InStr(1, Linha, "=", vbTextCompare) + 1)))
            
            ElseIf InStr(1, Linha, "CommandName", vbTextCompare) > 0 And DentroCon = False And DentroRS = True And DentroCampo = False And DentroPar = False Then
            NomeComando = Trim(Mid(Linha, InStr(1, Linha, "=", vbTextCompare) + 1))
            NomeComando = Mid(NomeComando, 2)
            NomeComando = Mid(NomeComando, 1, Len(NomeComando) - 1)
            
            
            xKey = NomeComando & "COMMAND:" & String(6 - Len(Trim(Str(InicioRS))), "0") & Trim(Str(InicioRS)) & "/"
            Nodes = TVOriginal.Nodes.Add(, , xKey, NomeComando, "Commando")
            NoRS = TVOriginal.Nodes.Count
            ElseIf InStr(1, Linha, "Name", vbTextCompare) > 0 And DentroCon = False And DentroRS = True And DentroCampo = True And DentroPar = False Then
            'verificao detalhada da palavra name. Ex. treiNAMEnto tem NAME porém não é a palavra chave procurada.
            'Para assegurar que o conteudo seja mesmo a palavra chave NAME, verifica-se primeiramente se
            'o trecho NAME existe na String. Caso exista, verifica-se se o sinal IGUAL(=) existe na String. Caso
            'exista, verifica-se se o trecho NAME está antes do sinal IGUAL(=). Caso esteja, verifica-se se antes do
            'sinal IGUAL(=) o único conteúdo que existe seja o trecho NAME.
            
                If InStr(1, Linha, "=", vbTextCompare) > 0 Then
                    If InStr(1, Linha, "Name", vbTextCompare) < InStr(1, Linha, "=", vbTextCompare) Then
                        If UCase(Trim(Mid(Linha, 1, InStr(1, Linha, "=", vbTextCompare) - 1))) = "NAME" Then
                        NomeCampo = Trim(Mid(Linha, InStr(1, Linha, "=", vbTextCompare) + 1))
                        
                        NomeCampo = Mid(NomeCampo, 2)
                        NomeCampo = Mid(NomeCampo, 1, Len(NomeCampo) - 1)
                        
                        Nodes = TVOriginal.Nodes.Add(xKey, tvwChild, NomeComando & "Field" & IndexCampo, NomeCampo, "Campo")
                        End If
                    End If
                End If
            End If
        
        FlexLinha.Rows = FlexLinha.Rows + 1
        FlexLinha.TextMatrix(FlexLinha.Rows - 1, Nivel - 1) = Linha
        End If
    Loop
Close #1

    For x = 0 To FlexLinha.Cols - 1
    FlexLinha.ColWidth(x) = 4000
    Next

FlexLinha.FixedCols = 0
FlexLinha.FixedRows = 0
CmdCarregar.Enabled = True
CmdOrdenar.Enabled = True
CmdGerar.Enabled = True
CmdSair.Enabled = True
Me.MousePointer = 0
LblMensagem.Caption = "Leitura Finalizada."
DoEvents
End Sub

Private Sub CmdGerar_Click()
Dim xLinhaIn As Long, xLinhaFin As Long, xIndexRS As Long, y As Long, xLinArqOrig As Long
Dim Linha As String
Dim xArquivoOriginal As String, xArquivoNovo As String, xNomeVerdadeiro As String
Dim CaminhoArquivoNovo As String, NomeArquivoNovo As String

    If TVOrdenada.Nodes.Count < 1 Then
    MsgBox "É necessário fazer a ordenação antes.", vbExclamation, ""
    Exit Sub
    End If


CDWin.DialogTitle = "Salvar Como"
CDWin.ShowSave

xArquivoNovo = CDWin.FileName

If xArquivoNovo = "" Then Exit Sub
    If UCase(Mid(xArquivoNovo, Len(xArquivoNovo) - 2)) <> "DSR" Then
    MsgBox "Tipo Inválido de Arquivo!", vbCritical, ""
    Exit Sub
    End If

xArquivoOriginal = LblArquivoOriginal.Caption

    If xArquivoOriginal = xArquivoNovo Then
        'If MsgBox("Deseja substituir o arquivo existente?", vbYesNo + vbExclamation, "") = vbYes Then
        'xNomeVerdadeiro = xArquivoNovo
        'xArquivoNovo = App.Path
        '    If Mid(xArquivoNovo, Len(xArquivoNovo), 1) <> "\" Then
        '    xArquivoNovo = xArquivoNovo & "\"
        '    End If
        'xArquivoNovo = xArquivoNovo & "TEMPCRI.DSR"
        'End If
    MsgBox "Por motivo de segurança, não é permitido salvar o novo arquivo com o mesmo nome do antigo. Salve o arquivo novo com um nome diferente e, despois de testá-lo, renomei-o para o nome desejado.", vbCritical, ""
    Exit Sub
    End If


CmdCarregar.Enabled = False
CmdOrdenar.Enabled = False
CmdGerar.Enabled = False
CmdSair.Enabled = False
Me.MousePointer = 11
LblMensagem.Caption = ""
DoEvents

Open xArquivoNovo For Output As #2

'Descobrindo ate em que linha vai o cabecalho

xLinhaIn = CDbl(Mid(TVOriginal.Nodes.Item(1).Key, (Len(TVOriginal.Nodes.Item(1).Key) - 20) + 8, 6)) - 1
Linha = ""

Open xArquivoOriginal For Input As #1

'Copiando Cabecalho
    For y = 1 To xLinhaIn
    Line Input #1, Linha
    Print #2, Linha
    Next

Close #1

xIndexRS = 0

    For y = 1 To TVOrdenada.Nodes.Count
        
    If TVOrdenada.Nodes.Item(y).FullPath = TVOrdenada.Nodes.Item(y).Text Then
    
    xIndexRS = xIndexRS + 1
    
    xLinhaIn = CDbl(Mid(TVOrdenada.Nodes.Item(y).Key, (Len(TVOrdenada.Nodes.Item(y).Key) - 20) + 8, 6))
    xLinhaFin = CDbl(Mid(TVOrdenada.Nodes.Item(y).Key, (Len(TVOrdenada.Nodes.Item(y).Key) - 5), 6))
    
    Open xArquivoOriginal For Input As #1
    
    xLinArqOrig = 0
    
    Linha = ""
        Do Until EOF(1)
        xLinArqOrig = xLinArqOrig + 1
        Line Input #1, Linha
            If xLinArqOrig >= xLinhaIn And xLinArqOrig <= xLinhaFin Then
                If InStr(1, Linha, "BeginProperty Recordset", vbTextCompare) > 0 Then
                Linha = RTrim(Linha)
                Linha = Mid(Linha, 1, InStr(1, Linha, "BeginProperty Recordset", vbTextCompare) + 22)
                Linha = Linha & xIndexRS
                End If
            Print #2, Linha
            ElseIf xLinArqOrig > xLinhaFin Then
            Exit Do
            End If
        Loop
    Close #1
    End If
    Next
    
Open xArquivoOriginal For Input As #1

xLinArqOrig = 0

Linha = ""
    Do Until EOF(1)
    xLinArqOrig = xLinArqOrig + 1
    Line Input #1, Linha
        If xLinArqOrig > xLinhaFin Then
        Print #2, Linha
        ElseIf xLinArqOrig > xLinhaFin Then
        Exit Do
        End If
    Loop
Close #1

Close #2


MsgBox "Arquivo Gravado com Sucesso!", vbInformation

CmdCarregar.Enabled = True
CmdOrdenar.Enabled = True
CmdGerar.Enabled = True
CmdSair.Enabled = True
Me.MousePointer = 0
LblMensagem.Caption = ""
DoEvents

End Sub

Private Sub CmdOrdenar_Click()
Dim xInsere As Boolean

    If TVOriginal.Nodes.Count < 1 Then
    MsgBox "É necessário carregar primeiros os comandos.", vbExclamation, ""
    Exit Sub
    End If

CmdCarregar.Enabled = False
CmdOrdenar.Enabled = False
CmdGerar.Enabled = False
CmdSair.Enabled = False
Me.MousePointer = 11
LblMensagem.Caption = ""
DoEvents

FlexOrdena.Clear
'FlexOrdena.FixedRows = 0
'FlexOrdena.FixedCols = 0
FlexOrdena.Rows = 0
FlexOrdena.Cols = 2

LblMensagem.Caption = "Lendo os nós de comando..."
ProgressBar.Max = TVOriginal.Nodes.Count
ProgressBar.Value = 0
DoEvents

For y = 1 To TVOriginal.Nodes.Count
    If InStr(1, TVOriginal.Nodes.Item(y).Key, "Field", vbTextCompare) = 0 Then
        If FlexOrdena.Rows = 0 Then
        FlexOrdena.Rows = 1
        FlexOrdena.Cols = 2
        Else
        FlexOrdena.Rows = FlexOrdena.Rows + 1
        End If
    
    FlexOrdena.TextMatrix(FlexOrdena.Rows - 1, 0) = TVOriginal.Nodes.Item(y).Text
    FlexOrdena.TextMatrix(FlexOrdena.Rows - 1, 1) = y
    End If
ProgressBar.Value = y
DoEvents
Next

FlexOrdena.ColWidth(0) = 5000
FlexOrdena.ColWidth(1) = 3000

LblMensagem.Caption = "Ordenando os comandos..."
ProgressBar.Max = FlexOrdena.Rows - 1
ProgressBar.Value = 0
DoEvents


    For i = 0 To FlexOrdena.Rows - 1
        For j = 0 To FlexOrdena.Rows - 2
            If UCase(FlexOrdena.TextMatrix(i, 0)) < UCase(FlexOrdena.TextMatrix(j, 0)) Then
            xAUX0 = FlexOrdena.TextMatrix(j, 0)
            xAUX1 = FlexOrdena.TextMatrix(j, 1)
            FlexOrdena.TextMatrix(j, 0) = FlexOrdena.TextMatrix(i, 0)
            FlexOrdena.TextMatrix(j, 1) = FlexOrdena.TextMatrix(i, 1)
            FlexOrdena.TextMatrix(i, 0) = xAUX0
            FlexOrdena.TextMatrix(i, 1) = xAUX1
            End If
        Next
    ProgressBar.Value = i
    DoEvents
    Next

TVOrdenada.Nodes.Clear
LblMensagem.Caption = "Transferindo nós organizados..."
ProgressBar.Max = FlexOrdena.Rows - 1
ProgressBar.Value = 0
DoEvents

    For x = 0 To FlexOrdena.Rows - 1
    xInsere = True
        For y = Val(FlexOrdena.TextMatrix(x, 1)) To TVOriginal.Nodes.Count
            If y = Val(FlexOrdena.TextMatrix(x, 1)) Then
            NomeComando = TVOriginal.Nodes.Item(y).Key
            Nodes = TVOrdenada.Nodes.Add(, , TVOriginal.Nodes.Item(y).Key, TVOriginal.Nodes.Item(y).Text, TVOriginal.Nodes.Item(y).Image)
            Else
            xInsere = True
                For a = 0 To FlexOrdena.Rows - 1
                    If Val(FlexOrdena.TextMatrix(a, 1)) = y Then
                    xInsere = False
                    Exit For
                    End If
                Next
            
                If xInsere = True Then
                Nodes = TVOrdenada.Nodes.Add(NomeComando, tvwChild, TVOriginal.Nodes.Item(y).Key, TVOriginal.Nodes.Item(y).Text, TVOriginal.Nodes.Item(y).Image)
                End If
            End If
            
            If xInsere = False Then
            Exit For
            End If
        Next
    ProgressBar.Value = x
    DoEvents
    Next

CmdCarregar.Enabled = True
CmdOrdenar.Enabled = True
CmdGerar.Enabled = True
CmdSair.Enabled = True
Me.MousePointer = 0
LblMensagem.Caption = ""
ProgressBar.Value = 0
DoEvents

End Sub

Private Sub CmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
'FlexLinha.FixedRows = 0
'FlexLinha.FixedCols = 0
FlexLinha.Rows = 0
End Sub
