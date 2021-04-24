VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReportCIA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório"
   ClientHeight    =   3570
   ClientLeft      =   1665
   ClientTop       =   3000
   ClientWidth     =   9390
   ControlBox      =   0   'False
   Icon            =   "frmReportCIA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraFiliais 
      Caption         =   "Informe a UF de Emissão dos AWBs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3735
      Begin VB.TextBox TxtBuscaFilial 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Top             =   300
         Width           =   435
      End
      Begin VB.Label TxtInscrEstFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   22
         Top             =   1140
         Width           =   2655
      End
      Begin VB.Label TxtCGCFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   21
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Est."
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1185
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   420
         TabIndex        =   19
         Top             =   885
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   330
         TabIndex        =   18
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label TxtUFFilial 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3060
         TabIndex        =   17
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Lic. IATA"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label TxtSiglaFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2940
         TabIndex        =   15
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label TxtNomeFilial 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label TxtFilial 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   300
         Width           =   435
      End
      Begin VB.Label TxtCidadeFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   12
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label TxtLicensaFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   11
         Top             =   1740
         Width           =   1995
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3540
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vencimento do Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   36
      Top             =   3720
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox ChkVER 
         Caption         =   "Visualizar Excel"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame FraArquivo 
      Caption         =   "Local de Gravação do Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   3900
      TabIndex        =   35
      Top             =   120
      Width           =   5355
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   2760
         TabIndex        =   6
         Top             =   660
         Width           =   2475
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   180
         Pattern         =   "*.xls"
         TabIndex        =   7
         Top             =   660
         Width           =   2475
      End
      Begin VB.TextBox TxtNomeArquivo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   180
         MaxLength       =   30
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   5700
      TabIndex        =   9
      Top             =   3060
      Width           =   1755
   End
   Begin VB.Frame FraDatas 
      Caption         =   "Período de Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   3735
      Begin MSMask.MaskEdBox MskDataInicial 
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDataFinal 
         Height          =   285
         Left            =   2220
         TabIndex        =   3
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   1920
         TabIndex        =   34
         Top             =   345
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   345
         Width           =   210
      End
   End
   Begin VB.Frame FraVencimento 
      Caption         =   "Vencimento do Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   30
      Top             =   2280
      Width           =   3735
      Begin MSMask.MaskEdBox MskVencimento 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Data do Vencimento"
         Height          =   195
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FraCiaAerea 
      Caption         =   "Cia. Aérea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   3735
      Begin VB.TextBox TxtSiglaCiaAerea 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox TxtBuscaSiglaCia 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   1
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Inscr. Est."
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   1185
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   420
         TabIndex        =   28
         Top             =   885
         Width           =   405
      End
      Begin VB.Label TxtCGCCiaAerea 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   27
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label TxtInscrEstCiaAerea 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   26
         Top             =   1140
         Width           =   2655
      End
      Begin VB.Label TxtNomeCiaAerea 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   300
         Width           =   2535
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3540
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   7500
      TabIndex        =   8
      Top             =   3060
      Width           =   1755
   End
End
Attribute VB_Name = "frmReportCIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

Dim xDataInicial As String
Dim xDataFinal As String
Dim xCia As String
Dim xFilial As String


    If Len(Trim(TxtFilial.Caption)) = 0 Then
    MsgBox "Você não informou a UF de emissão.", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(TxtSiglaCiaAerea.Text)) = 0 Then
    MsgBox "Você não informou a Cia. Aérea.", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(TxtNomeArquivo.Text)) = 0 Then
    MsgBox "Você não informou o nome do arquivo.", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(MskDataFinal.Text)) = 0 Or Len(Trim(MskDataInicial.Text)) = 0 Or Len(Trim(MskVencimento.Text)) = 0 Then
    MsgBox "Uma das datas necessárias não foi informada corretamente.", vbCritical, ""
    Exit Sub
    ElseIf CDate(MskDataFinal.Text) < CDate(MskDataInicial.Text) Then
    MsgBox "Sua data inicial é maior que a data final.", vbCritical, ""
    Exit Sub
    End If

cmdSair.Enabled = False
Command1.Enabled = False
FraFiliais.Enabled = False
FraCiaAerea.Enabled = False
FraArquivo.Enabled = False
FraDatas.Enabled = False
FraVencimento.Enabled = False

xDataInicial = CDate(MskDataInicial.Text)
xDataFinal = CDate(MskDataFinal.Text)

xCia = UCase(Trim(TxtSiglaCiaAerea.Text))
xFilial = Trim(TxtFilial.Caption)
    


    If de_informa.rsNúmerosAWBOK.State = 1 Then de_informa.rsNúmerosAWBOK.Close
    If de_informa.rsFreteAliquota.State = 1 Then de_informa.rsFreteAliquota.Close
    If de_informa.rsFreteModal.State = 1 Then de_informa.rsFreteModal.Close
    If de_informa.rsFreteTipoADVAL.State = 1 Then de_informa.rsFreteTipoADVAL.Close
    If de_informa.rsNúmerosAWBCANC.State = 1 Then de_informa.rsNúmerosAWBCANC.Close
    
    de_informa.NúmerosAWBOK CDate(xDataInicial), CDate(xDataFinal), xCia, xFilial
    de_informa.FreteAliquota CDate(xDataInicial), CDate(xDataFinal), xCia, xFilial
    de_informa.FreteModal CDate(xDataInicial), CDate(xDataFinal), xCia, xFilial
    de_informa.FreteTipoADVAL CDate(xDataInicial), CDate(xDataFinal), xCia, xFilial
    de_informa.NúmerosAWBCANC CDate(xDataInicial), CDate(xDataFinal), xCia, xFilial
    
    If de_informa.rsNúmerosAWBOK.RecordCount = 0 Then
    MsgBox "Para esta pesquisa não foram encontrados AWBs. Por favor, tente novemente.", vbCritical, ""
    Exit Sub
    End If
    
Dim Excel As Excel.Application
Dim ExcelWBK As Excel.Workbook
Dim ExcelWS As Excel.Worksheet
Dim ExcelWS2 As Excel.Worksheet
Dim ExcelWS3 As Excel.Worksheet
Dim ExcelWS4 As Excel.Worksheet
Dim ExcelA1 As Excel.Worksheet
Dim ExcelA2 As Excel.Worksheet
Dim ExcelA3 As Excel.Worksheet

Set Excel = CreateObject("EXCEL.APPLICATION")
Excel.Visible = True
Excel.Interactive = True
Excel.DisplayAlerts = False

Excel.Workbooks.Add

Set ExcelWBK = Excel.ActiveWorkbook

Set ExcelWS = Excel.ActiveSheet

Dim Linha As Integer
Linha = 1

ExcelWS.Cells(1, 1) = "Relatório de Emissão de Conhecimentos Aéreos"
ExcelWS.Cells(2, 1) = "Agente"
ExcelWS.Cells(3, 1) = "Localidade"
ExcelWS.Cells(4, 1) = "Período de Vendas"
ExcelWS.Cells(5, 1) = "Cia. Aérea"
ExcelWS.Cells(6, 1) = "Emissor"
ExcelWS.Cells(7, 1) = "Cod. IATA"
'ExcelWS.Cells(7, 2) = "Compl."
'ExcelWS.Cells(7, 3) = "CK"
ExcelWS.Cells(6, 6) = "Cod. Localidade"
ExcelWS.Cells(7, 6) = "Alfa"
ExcelWS.Cells(7, 7) = "Numer."
ExcelWS.Cells(7, 8) = "Qtde."
ExcelWS.Cells(7, 9) = "Tipo"
ExcelWS.Cells(10, 1) = "Número de Série dos Conhecimentos Aéreos"
ExcelWS.Cells(11, 1) = "Formulário"
ExcelWS.Cells(11, 2) = "Qtde."
ExcelWS.Cells(11, 3) = "De"
ExcelWS.Cells(11, 4) = "CK"
ExcelWS.Cells(11, 5) = "Até"
ExcelWS.Cells(11, 6) = "CK"

ExcelWS.Cells(21, 2) = "Frete Nacional"
ExcelWS.Cells(21, 3) = "Frete Regional"
ExcelWS.Cells(21, 4) = "Taxas"
ExcelWS.Cells(21, 5) = "Ad. Val."
ExcelWS.Cells(21, 6) = "Total"
    
ExcelWS.Cells(21, 1) = "Cod. Fisc."
ExcelWS.Cells(22, 1) = "12%"
ExcelWS.Cells(23, 1) = "4%"
ExcelWS.Cells(24, 1) = "ISENTO"

ExcelWS.Cells(27, 1) = "Forma de Pagamento"
ExcelWS.Cells(27, 3) = "Tipo"
ExcelWS.Cells(27, 4) = "Qtde."
ExcelWS.Cells(27, 5) = "Frete Nac."
ExcelWS.Cells(27, 6) = "Frete Reg."
ExcelWS.Cells(27, 7) = "Tx. Terr."
ExcelWS.Cells(27, 8) = "Ad. Val."
ExcelWS.Cells(27, 9) = "Frete a Cobrar (FRAP)"
ExcelWS.Cells(27, 10) = "Local"
ExcelWS.Cells(27, 11) = "Não Local"
ExcelWS.Cells(27, 12) = "Caixa"
ExcelWS.Cells(26, 10) = "Frete Pago"
ExcelWS.Cells(28, 1) = "Modalidade"
ExcelWS.Cells(28, 2) = "A Cobrar"
ExcelWS.Cells(29, 2) = "C. Corrente"
ExcelWS.Cells(30, 2) = "Pago"
ExcelWS.Cells(31, 2) = "COD"
ExcelWS.Cells(28, 3) = "4"
ExcelWS.Cells(29, 3) = "5"
ExcelWS.Cells(30, 3) = "6"
ExcelWS.Cells(31, 3) = "9"
ExcelWS.Cells(32, 1) = "Totais"

ExcelWS.Cells(34, 1) = "Ad. Valorem"
ExcelWS.Cells(34, 2) = "Cod."
ExcelWS.Cells(34, 3) = "Valor"
ExcelWS.Cells(35, 1) = "Carga Comum"
ExcelWS.Cells(36, 1) = "Carga Valor"
ExcelWS.Cells(37, 1) = "Animais e Aves Vivas"
ExcelWS.Cells(35, 2) = "1"
ExcelWS.Cells(36, 2) = "2"
ExcelWS.Cells(37, 2) = "3"

ExcelWS.Cells(39, 1) = "Nº de FRAP Devolvidos/Redespachados"
ExcelWS.Cells(40, 1) = "Nº de COD Devolvidos"
ExcelWS.Cells(41, 1) = "Total de Frete a Cobrar no Destino"

ExcelWS.Cells(43, 1) = "Comissão 6% sobre Frete (-)"
ExcelWS.Cells(44, 1) = "Recolhimento IR sobre Comissão"
ExcelWS.Cells(45, 1) = "Taxa Devida ao Agente na Origem"

ExcelWS.Cells(47, 1) = "Vl. a Reemb. ao Rem."
ExcelWS.Cells(47, 3) = "Tx Orig,, Dest. e Red."
ExcelWS.Cells(47, 5) = "Tx Serv COD"

ExcelWS.Cells(49, 1) = "Total Líquido"

ExcelWS.Cells(51, 1) = "Vencimento"

ExcelWS.Cells(53, 1) = "Comissão 12%"
ExcelWS.Cells(54, 1) = "Comissão 4%"
ExcelWS.Cells(55, 1) = "Comissão ISENTO"

ExcelWS.Cells(57, 1) = "Data"

    
    
ExcelWS.Cells(2, 4) = "INTEC Transportes LTDA."
ExcelWS.Cells(3, 4) = "São Paulo"
ExcelWS.Cells(4, 4) = xDataInicial & " até " & xDataFinal
ExcelWS.Cells(5, 4) = PriMaiuscula(TxtNomeCiaAerea.Caption)
ExcelWS.Cells(8, 1) = "57060640011"
ExcelWS.Cells(8, 2) = "001"
ExcelWS.Cells(8, 3) = "6"
ExcelWS.Cells(8, 6) = "SAO"
ExcelWS.Cells(8, 7) = "1912"
ExcelWS.Cells(8, 8) = "98"
ExcelWS.Cells(8, 9) = "1"

Dim xFrmI01 As String
Dim xFrmF01 As String
Dim xDigI01 As String
Dim xDigF01 As String
Dim xQdte01 As Integer

Dim xFrmI02 As String
Dim xFrmF02 As String
Dim xDigI02 As String
Dim xDigF02 As String
Dim xQdte02 As Integer

Dim xFrmI03 As String
Dim xFrmF03 As String
Dim xDigI03 As String
Dim xDigF03 As String
Dim xQdte03 As Integer

Dim xFrmI04 As String
Dim xFrmF04 As String
Dim xDigI04 As String
Dim xDigF04 As String
Dim xQdte04 As Integer

Dim xFrmI05 As String
Dim xFrmF05 As String
Dim xDigI05 As String
Dim xDigF05 As String
Dim xQdte05 As Integer

Dim xFrmI06 As String
Dim xFrmF06 As String
Dim xDigI06 As String
Dim xDigF06 As String
Dim xQdte06 As Integer

Dim xFrmI07 As String
Dim xFrmF07 As String
Dim xDigI07 As String
Dim xDigF07 As String
Dim xQdte07 As Integer

Dim xFrmI08 As String
Dim xFrmF08 As String
Dim xDigI08 As String
Dim xDigF08 As String
Dim xQdte08 As Integer

Dim xFrmI09 As String
Dim xFrmF09 As String
Dim xDigI09 As String
Dim xDigF09 As String
Dim xQdte09 As Integer

Dim xFrmI10 As String
Dim xFrmF10 As String
Dim xDigI10 As String
Dim xDigF10 As String
Dim xQdte10 As Integer

xQtde01 = 0
xQtde02 = 0
xQtde03 = 0
xQtde04 = 0
xQtde05 = 0
xQtde06 = 0
xQtde07 = 0
xQtde08 = 0
xQtde09 = 0
xQtde10 = 0

Dim Seq As Integer

Dim xFrmI As String
Dim xFrmF As String
Dim xDigI As String
Dim xDigF As String
Dim xQdte As Integer

Dim xFProx As String

Dim xUltimo As Boolean

Seq = 1

xFormI = ""
xFormF = "  "
xDigI = ""
xDigF = ""
xQtde = 0

    
    Do Until de_informa.rsNúmerosAWBOK.EOF
    xQtde = xQtde + 1
        If xQtde = 1 Then
        xFormI = de_informa.rsNúmerosAWBOK.Fields("awb")
        xDigI = de_informa.rsNúmerosAWBOK.Fields("dig")
        Else
        xFormF = de_informa.rsNúmerosAWBOK.Fields("awb")
        xDigF = de_informa.rsNúmerosAWBOK.Fields("dig")
        End If
        
    
    de_informa.rsNúmerosAWBOK.MoveNext
        If de_informa.rsNúmerosAWBOK.EOF = False Then
        xFProx = de_informa.rsNúmerosAWBOK.Fields("awb")
        Else
        xUltimo = True
        xFProx = xFormF
        End If
    de_informa.rsNúmerosAWBOK.MovePrevious
    
        
        If (Mid(xFormF, Len(xFormF) - 1) = "00" And xQtde > 1) Or ((Val(xFProx) - Val(xFormF)) > 10 And Val(xFormF) > 0) Or xUltimo = True Then

            If Seq = 1 Then
            xFrmI01 = xFormI
            xFrmF01 = xFormF
            xDigI01 = xDigI
            xDigF01 = xDigF
            xQdte01 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 2 Then
            xFrmI02 = xFormI
            xFrmF02 = xFormF
            xDigI02 = xDigI
            xDigF02 = xDigF
            xQdte02 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 3 Then
            xFrmI03 = xFormI
            xFrmF03 = xFormF
            xDigI03 = xDigI
            xDigF03 = xDigF
            xQdte03 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 4 Then
            xFrmI04 = xFormI
            xFrmF04 = xFormF
            xDigI04 = xDigI
            xDigF04 = xDigF
            xQdte04 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 5 Then
            xFrmI05 = xFormI
            xFrmF05 = xFormF
            xDigI05 = xDigI
            xDigF05 = xDigF
            xQdte05 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 6 Then
            xFrmI06 = xFormI
            xFrmF06 = xFormF
            xDigI06 = xDigI
            xDigF06 = xDigF
            xQdte06 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 7 Then
            xFrmI07 = xFormI
            xFrmF07 = xFormF
            xDigI07 = xDigI
            xDigF07 = xDigF
            xQdte07 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 8 Then
            xFrmI08 = xFormI
            xFrmF08 = xFormF
            xDigI08 = xDigI
            xDigF08 = xDigF
            xQdte08 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 9 Then
            xFrmI09 = xFormI
            xFrmF09 = xFormF
            xDigI09 = xDigI
            xDigF09 = xDigF
            xQdte09 = xQtde
            Seq = Seq + 1
            ElseIf Seq = 10 Then
            xFrmI10 = xFormI
            xFrmF10 = xFormF
            xDigI10 = xDigI
            xDigF10 = xDigF
            xQdte10 = xQtde
            Seq = Seq + 1
            End If
        xFormI = ""
        xFormF = "  "
        xDigI = ""
        xDigF = ""
        xQtde = 0
        End If
  
    de_informa.rsNúmerosAWBOK.MoveNext
    Loop

ExcelWS.Cells(12, 2) = xQdte01
If xQdte01 = 0 Then ExcelWS.Cells(12, 2) = ""
ExcelWS.Cells(12, 3) = xFrmI01
ExcelWS.Cells(12, 4) = xDigI01
ExcelWS.Cells(12, 5) = xFrmF01
ExcelWS.Cells(12, 6) = xDigF01

ExcelWS.Cells(13, 2) = xQdte02
If xQdte02 = 0 Then ExcelWS.Cells(13, 2) = ""
ExcelWS.Cells(13, 3) = xFrmI02
ExcelWS.Cells(13, 4) = xDigI02
ExcelWS.Cells(13, 5) = xFrmF02
ExcelWS.Cells(13, 6) = xDigF02

ExcelWS.Cells(14, 2) = xQdte03
If xQdte03 = 0 Then ExcelWS.Cells(14, 2) = ""
ExcelWS.Cells(14, 3) = xFrmI03
ExcelWS.Cells(14, 4) = xDigI03
ExcelWS.Cells(14, 5) = xFrmF03
ExcelWS.Cells(14, 6) = xDigF03

ExcelWS.Cells(15, 2) = xQdte04
If xQdte04 = 0 Then ExcelWS.Cells(15, 2) = ""
ExcelWS.Cells(15, 3) = xFrmI04
ExcelWS.Cells(15, 4) = xDigI04
ExcelWS.Cells(15, 5) = xFrmF04
ExcelWS.Cells(15, 6) = xDigF04

ExcelWS.Cells(16, 2) = xQdte05
If xQdte05 = 0 Then ExcelWS.Cells(16, 2) = ""
ExcelWS.Cells(16, 3) = xFrmI05
ExcelWS.Cells(16, 4) = xDigI05
ExcelWS.Cells(16, 5) = xFrmF05
ExcelWS.Cells(16, 6) = xDigF05

ExcelWS.Cells(17, 2) = xQdte06
If xQdte06 = 0 Then ExcelWS.Cells(17, 2) = ""
ExcelWS.Cells(17, 3) = xFrmI06
ExcelWS.Cells(17, 4) = xDigI06
ExcelWS.Cells(17, 5) = xFrmF06
ExcelWS.Cells(17, 6) = xDigF06

ExcelWS.Cells(18, 2) = xQdte07
If xQdte07 = 0 Then ExcelWS.Cells(18, 2) = ""
ExcelWS.Cells(18, 3) = xFrmI07
ExcelWS.Cells(18, 4) = xDigI07
ExcelWS.Cells(18, 5) = xFrmF07
ExcelWS.Cells(18, 6) = xDigF07


Dim xFreteN12 As Double
Dim xFreteR12 As Double
Dim xTaxas12 As Double
Dim xADVAL12 As Double
Dim xTotal12 As Double

Dim xFreteN4 As Double
Dim xFreteR4 As Double
Dim xTaxas4 As Double
Dim xADVAL4 As Double
Dim xTotal4 As Double

Dim xFreteNI As Double
Dim xFreteRI As Double
Dim xTaxasI As Double
Dim xADVALI As Double
Dim xTotalI As Double

    Do Until de_informa.rsFreteAliquota.EOF
        If de_informa.rsFreteAliquota.Fields("aliquota") = "4" Then
        xFreteN4 = de_informa.rsFreteAliquota.Fields("fretenacional")
        xFreteR4 = de_informa.rsFreteAliquota.Fields("freteregional")
        xTaxas4 = de_informa.rsFreteAliquota.Fields("taxas")
        xADVAL4 = de_informa.rsFreteAliquota.Fields("advalorem")
        xTotal4 = de_informa.rsFreteAliquota.Fields("fretetotal")
        ElseIf de_informa.rsFreteAliquota.Fields("aliquota") = "12" Then
        xFreteN12 = de_informa.rsFreteAliquota.Fields("fretenacional")
        xFreteR12 = de_informa.rsFreteAliquota.Fields("freteregional")
        xTaxas12 = de_informa.rsFreteAliquota.Fields("taxas")
        xADVAL12 = de_informa.rsFreteAliquota.Fields("advalorem")
        xTotal12 = de_informa.rsFreteAliquota.Fields("fretetotal")
        ElseIf de_informa.rsFreteAliquota.Fields("aliquota") = "ISENTO" Then
        xFreteNI = de_informa.rsFreteAliquota.Fields("fretenacional")
        xFreteRI = de_informa.rsFreteAliquota.Fields("freteregional")
        xTaxasI = de_informa.rsFreteAliquota.Fields("taxas")
        xADVALI = de_informa.rsFreteAliquota.Fields("advalorem")
        xTotalI = de_informa.rsFreteAliquota.Fields("fretetotal")
        End If
    de_informa.rsFreteAliquota.MoveNext
    Loop


ExcelWS.Cells(22, 2) = xFreteN12
ExcelWS.Cells(22, 3) = xFreteR12
ExcelWS.Cells(22, 4) = xTaxas12
ExcelWS.Cells(22, 5) = xADVAL12
ExcelWS.Cells(22, 6) = xTotal12

ExcelWS.Cells(23, 2) = xFreteN4
ExcelWS.Cells(23, 3) = xFreteR4
ExcelWS.Cells(23, 4) = xTaxas4
ExcelWS.Cells(23, 5) = xADVAL4
ExcelWS.Cells(23, 6) = xTotal4

ExcelWS.Cells(24, 2) = xFreteNI
ExcelWS.Cells(24, 3) = xFreteRI
ExcelWS.Cells(24, 4) = xTaxasI
ExcelWS.Cells(24, 5) = xADVALI
ExcelWS.Cells(24, 6) = xTotalI


Dim xFreteNAP As Double
Dim xFreteRAP As Double
Dim xTaxasAP As Double
Dim xAdValAP As Double
Dim xTotalAP As Double
Dim xQtdeAP As Double

Dim xFreteNP As Double
Dim xFreteRP As Double
Dim xTaxasP As Double
Dim xAdValP As Double
Dim xTotalP As Double
Dim xQtdeP As Double

    Do Until de_informa.rsFreteModal.EOF
        If de_informa.rsFreteModal.Fields("modal") = "A PAGAR" Then
        xFreteNAP = de_informa.rsFreteModal.Fields("fretenacional")
        xFreteRAP = de_informa.rsFreteModal.Fields("freteregional")
        xTaxasAP = de_informa.rsFreteModal.Fields("taxas")
        xAdValAP = de_informa.rsFreteModal.Fields("advalorem")
        xTotalAP = de_informa.rsFreteModal.Fields("fretetotal")
        xQtdeAP = de_informa.rsFreteModal.Fields("qtde")
        ElseIf de_informa.rsFreteModal.Fields("modal") = "PAGO" Then
        xFreteNP = de_informa.rsFreteModal.Fields("fretenacional")
        xFreteRP = de_informa.rsFreteModal.Fields("freteregional")
        xTaxasP = de_informa.rsFreteModal.Fields("taxas")
        xAdValP = de_informa.rsFreteModal.Fields("advalorem")
        xTotalP = de_informa.rsFreteModal.Fields("fretetotal")
        xQtdeP = de_informa.rsFreteModal.Fields("qtde")
        End If
    de_informa.rsFreteModal.MoveNext
    Loop

ExcelWS.Cells(28, 4) = xQtdeAP
ExcelWS.Cells(28, 5) = xFreteNAP
ExcelWS.Cells(28, 6) = xFreteRAP
ExcelWS.Cells(28, 7) = xTaxasAP
ExcelWS.Cells(28, 8) = xAdValAP
ExcelWS.Cells(28, 9) = xTotalAP

ExcelWS.Cells(30, 4) = xQtdeP
ExcelWS.Cells(30, 5) = xFreteNP
ExcelWS.Cells(30, 6) = xFreteRP
ExcelWS.Cells(30, 7) = xTaxasP
ExcelWS.Cells(30, 8) = xAdValP
ExcelWS.Cells(30, 12) = xTotalP

ExcelWS.Cells(32, 4) = xQtdeAP + xQtdeP
ExcelWS.Cells(32, 5) = xFreteNAP + xFreteNP
ExcelWS.Cells(32, 6) = xFreteRAP + xFreteRP
ExcelWS.Cells(32, 7) = xTaxasAP + xTaxasP
ExcelWS.Cells(32, 8) = xAdValAP + xAdValP
ExcelWS.Cells(32, 9) = xTotalAP
ExcelWS.Cells(32, 12) = xTotalP


Dim Adval1 As Double
Dim Adval2 As Double
Dim Adval3 As Double

    Do Until de_informa.rsFreteTipoADVAL.EOF
        If de_informa.rsFreteTipoADVAL.Fields("tipoadval") = "1" Then
        Adval1 = de_informa.rsFreteTipoADVAL.Fields("advalorem")
        ElseIf de_informa.rsFreteTipoADVAL.Fields("tipoadval") = "2" Then
        Adval2 = de_informa.rsFreteTipoADVAL.Fields("advalorem")
        ElseIf de_informa.rsFreteTipoADVAL.Fields("tipoadval") = "3" Then
        Adval3 = de_informa.rsFreteTipoADVAL.Fields("advalorem")
        End If
    de_informa.rsFreteTipoADVAL.MoveNext
    Loop

ExcelWS.Cells(35, 3) = Adval1
ExcelWS.Cells(36, 3) = Adval2
ExcelWS.Cells(37, 3) = Adval3


ExcelWS.Cells(39, 4) = ""
ExcelWS.Cells(40, 4) = ""
ExcelWS.Cells(41, 4) = xTotalAP

xperccom = 6
xpercir = 1.5
xcomissao = (((xFreteNAP + xFreteNP) / 100) * xperccom)
xir = (xcomissao / 100) * xpercir

ExcelWS.Cells(43, 4) = xcomissao
ExcelWS.Cells(44, 4) = xir
ExcelWS.Cells(45, 4) = ""

ExcelWS.Cells(47, 2) = ""
ExcelWS.Cells(47, 4) = xTaxas4 + xTaxas12 + xTaxasI
ExcelWS.Cells(47, 6) = ""

ExcelWS.Cells(49, 4) = (xTotalAP + xTotalP) - xcomissao

ExcelWS.Cells(51, 4) = MskVencimento.Text

ExcelWS.Cells(53, 4) = (xFreteN12 / 100) * xperccom
ExcelWS.Cells(54, 4) = (xFreteN4 / 100) * xperccom
ExcelWS.Cells(55, 4) = (xFreteNI / 100) * xperccom

If Month(Date) = 1 Then xmes = "Janeiro"
If Month(Date) = 2 Then xmes = "Fevereiro"
If Month(Date) = 3 Then xmes = "Março"
If Month(Date) = 4 Then xmes = "Abril"
If Month(Date) = 5 Then xmes = "Maio"
If Month(Date) = 6 Then xmes = "Junho"
If Month(Date) = 7 Then xmes = "Julho"
If Month(Date) = 8 Then xmes = "Agosto"
If Month(Date) = 9 Then xmes = "Setembro"
If Month(Date) = 10 Then xmes = "Outubro"
If Month(Date) = 11 Then xmes = "Novembro"
If Month(Date) = 12 Then xmes = "Dezembro"

xcancelados = ""
    Do Until de_informa.rsNúmerosAWBCANC.EOF
    xcancelados = xcancelados & de_informa.rsNúmerosAWBCANC.Fields("awb") & "-" & de_informa.rsNúmerosAWBCANC.Fields("dig") & "/"
    de_informa.rsNúmerosAWBCANC.MoveNext
    Loop

xcancelados = Mid(xcancelados, 1, 500) & String(500 - Len(Mid(xcancelados, 1, 500)), " ")

ExcelWS.Cells(57, 1) = "AWBs Cancelados"
ExcelWS.Cells(58, 1) = Mid(xcancelados, 1, 100)
ExcelWS.Cells(59, 1) = Mid(xcancelados, 101, 100)
ExcelWS.Cells(60, 1) = Mid(xcancelados, 201, 100)
ExcelWS.Cells(61, 1) = Mid(xcancelados, 301, 100)
ExcelWS.Cells(62, 1) = Mid(xcancelados, 401, 100)

ExcelWS.Cells(64, 12) = "São Paulo, " & Day(Date) & " de " & xmes & " de " & Year(Date)
    
    
'FORMATACAO DA PLANILHA

ExcelWS.Cells.Font.Name = "Arial"
ExcelWS.Cells.Font.Size = 8

ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(1, 1)).Font.Name = "Arial"
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(1, 1)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(1, 1)).Font.Size = 12
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(1, 1)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(1, 1)).Font.ColorIndex = 3
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(1, 12)).Merge

ExcelWS.Range(ExcelWS.Cells(5, 4), ExcelWS.Cells(5, 4)).Font.Name = "Arial"
ExcelWS.Range(ExcelWS.Cells(5, 4), ExcelWS.Cells(5, 4)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(5, 4), ExcelWS.Cells(5, 4)).Font.Size = 12
ExcelWS.Range(ExcelWS.Cells(5, 4), ExcelWS.Cells(5, 4)).Font.ColorIndex = 5

ExcelWS.Range(ExcelWS.Cells(2, 1), ExcelWS.Cells(2, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(3, 1), ExcelWS.Cells(3, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(4, 1), ExcelWS.Cells(4, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(5, 1), ExcelWS.Cells(5, 3)).Merge


ExcelWS.Range(ExcelWS.Cells(2, 4), ExcelWS.Cells(2, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(3, 4), ExcelWS.Cells(3, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(4, 4), ExcelWS.Cells(4, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(5, 4), ExcelWS.Cells(5, 12)).Merge

ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(6, 3)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(6, 3)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(6, 3)).Merge

ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(6, 9)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(6, 9)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(6, 9)).Merge

ExcelWS.Range(ExcelWS.Cells(7, 1), ExcelWS.Cells(7, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(8, 1), ExcelWS.Cells(8, 3)).Merge

ExcelWS.Range(ExcelWS.Cells(10, 1), ExcelWS.Cells(10, 1)).Font.Bold = True


ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 1)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 1)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 1)).VerticalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 1)).Merge


ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(21, 6)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(21, 6)).HorizontalAlignment = xlCenter

ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 1)).Font.Bold = True

ExcelWS.Range(ExcelWS.Cells(22, 2), ExcelWS.Cells(24, 6)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(27, 12)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(27, 12)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(27, 2)).Merge

ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).Merge

ExcelWS.Range(ExcelWS.Cells(28, 1), ExcelWS.Cells(31, 3)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(28, 1), ExcelWS.Cells(31, 1)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(28, 1), ExcelWS.Cells(31, 1)).VerticalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(28, 1), ExcelWS.Cells(31, 1)).Merge

ExcelWS.Range(ExcelWS.Cells(32, 1), ExcelWS.Cells(32, 12)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(32, 1), ExcelWS.Cells(32, 3)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(32, 1), ExcelWS.Cells(32, 3)).Merge

ExcelWS.Range(ExcelWS.Cells(28, 5), ExcelWS.Cells(32, 12)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(34, 3)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(35, 3), ExcelWS.Cells(37, 3)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(39, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(40, 1), ExcelWS.Cells(40, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(41, 1), ExcelWS.Cells(41, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(41, 4), ExcelWS.Cells(41, 4)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(43, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(44, 1), ExcelWS.Cells(44, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(45, 1), ExcelWS.Cells(45, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(43, 4), ExcelWS.Cells(45, 4)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(47, 2), ExcelWS.Cells(47, 2)).Style = "Comma"
ExcelWS.Range(ExcelWS.Cells(47, 4), ExcelWS.Cells(47, 4)).Style = "Comma"
ExcelWS.Range(ExcelWS.Cells(47, 6), ExcelWS.Cells(47, 6)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 4)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 4)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(49, 4), ExcelWS.Cells(49, 4)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 4)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 4)).HorizontalAlignment = xlCenter
ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 3)).Merge

ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(55, 4)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(53, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(54, 1), ExcelWS.Cells(54, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(55, 1), ExcelWS.Cells(55, 3)).Merge
ExcelWS.Range(ExcelWS.Cells(53, 4), ExcelWS.Cells(55, 4)).Style = "Comma"

ExcelWS.Range(ExcelWS.Cells(57, 1), ExcelWS.Cells(57, 4)).Font.Bold = True

ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Font.Name = "Courier New"
ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(58, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(59, 1), ExcelWS.Cells(59, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(60, 1), ExcelWS.Cells(60, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(61, 1), ExcelWS.Cells(61, 12)).Merge
ExcelWS.Range(ExcelWS.Cells(62, 1), ExcelWS.Cells(62, 12)).Merge

ExcelWS.Range(ExcelWS.Cells(57, 12), ExcelWS.Cells(57, 12)).Font.Bold = True
ExcelWS.Range(ExcelWS.Cells(57, 12), ExcelWS.Cells(57, 12)).HorizontalAlignment = xlRight

ExcelWS.Range(ExcelWS.Cells(57, 1), ExcelWS.Cells(57, 12)).Merge

ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(5, 12)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(5, 12)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(5, 12)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(5, 12)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(5, 12)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(1, 1), ExcelWS.Cells(5, 12)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(8, 3)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(8, 3)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(8, 3)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(8, 3)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(8, 3)).Borders(xlInsideHorizontal).Weight = xlThin
'ExcelWS.Range(ExcelWS.Cells(6, 1), ExcelWS.Cells(8, 3)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(8, 9)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(8, 9)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(8, 9)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(8, 9)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(8, 9)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(6, 6), ExcelWS.Cells(8, 9)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 6)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 6)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 6)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 6)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 6)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(11, 1), ExcelWS.Cells(18, 6)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 6)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 6)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 6)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 6)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 6)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(21, 1), ExcelWS.Cells(24, 6)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(32, 12)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(32, 12)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(32, 12)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(32, 12)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(32, 12)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(27, 1), ExcelWS.Cells(32, 12)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(26, 10), ExcelWS.Cells(26, 12)).Borders(xlEdgeTop).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(34, 1), ExcelWS.Cells(37, 3)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(39, 1), ExcelWS.Cells(41, 4)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(43, 1), ExcelWS.Cells(45, 4)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(47, 1), ExcelWS.Cells(47, 6)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(47, 1), ExcelWS.Cells(47, 6)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(47, 1), ExcelWS.Cells(47, 6)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(47, 1), ExcelWS.Cells(47, 6)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(47, 1), ExcelWS.Cells(47, 6)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 4)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 4)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 4)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(49, 1), ExcelWS.Cells(49, 4)).Borders(xlEdgeTop).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 4)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 4)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 4)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(51, 1), ExcelWS.Cells(51, 4)).Borders(xlEdgeTop).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(55, 4)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(55, 4)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(55, 4)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(55, 4)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(55, 4)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(53, 1), ExcelWS.Cells(45, 4)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Borders(xlEdgeBottom).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Borders(xlEdgeLeft).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Borders(xlEdgeRight).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Borders(xlEdgeTop).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Borders(xlInsideHorizontal).Weight = xlThin
ExcelWS.Range(ExcelWS.Cells(58, 1), ExcelWS.Cells(62, 12)).Borders(xlInsideVertical).Weight = xlThin

ExcelWS.Range(ExcelWS.Cells(1, 2), ExcelWS.Cells(62, 12)).EntireColumn.AutoFit

Excel.DisplayAlerts = False
        If Dir1.Path = "c:\" Then
        Excel.ActiveWorkbook.SaveAs Dir1.Path & PriMaiuscula((TxtNomeArquivo.Text)) & ".xls"
        Else
        Excel.ActiveWorkbook.SaveAs Dir1.Path & "\" & PriMaiuscula((TxtNomeArquivo.Text)) & ".xls"
        End If
Excel.DisplayAlerts = True

'Excel.Quit

Set ExcelWBK = Nothing
Set Excel = Nothing

Call LimpaTela(Me)

cmdSair.Enabled = True
Command1.Enabled = True
FraFiliais.Enabled = True
FraCiaAerea.Enabled = True
FraArquivo.Enabled = True
FraDatas.Enabled = True
FraVencimento.Enabled = True

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub File1_Click()
TxtNomeArquivo.Text = Mid(File1.FileName, 1, Len(File1.FileName) - 4)
End Sub

Private Sub Form_Load()

Dir1.Path = "c:\"
Dir1.Refresh
File1.Path = "c:\"
File1.Refresh

End Sub

Private Sub MskDataFinal_GotFocus()
Call Date_MskEdBox_GotFocus(MskDataFinal)
End Sub

Private Sub MskDataFinal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskDataFinal_LostFocus()
Call Date_MskEdBox_LostFocus(MskDataFinal)
End Sub

Private Sub MskDataInicial_GotFocus()
Call Date_MskEdBox_GotFocus(MskDataInicial)
End Sub

Private Sub MskDataInicial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskDataInicial_LostFocus()
Call Date_MskEdBox_LostFocus(MskDataInicial)
End Sub

Private Sub MskVencimento_GotFocus()
Call Date_MskEdBox_GotFocus(MskVencimento)
End Sub

Private Sub MskVencimento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub MskVencimento_LostFocus()
Call Date_MskEdBox_LostFocus(MskVencimento)
End Sub

Private Sub TxtBuscaFilial_Change()
X = TxtBuscaFilial.SelStart
TxtBuscaFilial.Text = UCase(TxtBuscaFilial.Text)
TxtBuscaFilial.SelStart = X
End Sub

Private Sub TxtBuscaFilial_GotFocus()
TxtBuscaFilial.SelStart = 0
TxtBuscaFilial.SelLength = 100
End Sub

Private Sub TxtBuscaFilial_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
End Sub

Private Sub TxtBuscaFilial_LostFocus()
TxtFilial.Caption = ""
TxtNomeFilial.Caption = ""

If TxtBuscaFilial.Text = "AC" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "AL" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "AM" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "AP" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "BA" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "CE" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "DF" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "ES" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "GO" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "MA" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "MG" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "MS" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "MT" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "PA" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "PB" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "PE" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "PI" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "PR" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "RJ" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "RN" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "RO" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "RR" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "RS" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "SC" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "SE" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "SP" Then TxtFilial.Caption = TxtBuscaFilial.Text
If TxtBuscaFilial.Text = "TO" Then TxtFilial.Caption = TxtBuscaFilial.Text

If TxtBuscaFilial.Text = "AC" Then TxtNomeFilial.Caption = "ACRE"
If TxtBuscaFilial.Text = "AL" Then TxtNomeFilial.Caption = "ALAGOAS"
If TxtBuscaFilial.Text = "AM" Then TxtNomeFilial.Caption = "AMAZONAS"
If TxtBuscaFilial.Text = "AP" Then TxtNomeFilial.Caption = "AMAPA"
If TxtBuscaFilial.Text = "BA" Then TxtNomeFilial.Caption = "BAHIA"
If TxtBuscaFilial.Text = "CE" Then TxtNomeFilial.Caption = "CEARA"
If TxtBuscaFilial.Text = "DF" Then TxtNomeFilial.Caption = "DISTRITO FEDERAL"
If TxtBuscaFilial.Text = "ES" Then TxtNomeFilial.Caption = "ESPIRITO SANTO"
If TxtBuscaFilial.Text = "GO" Then TxtNomeFilial.Caption = "GOIAS"
If TxtBuscaFilial.Text = "MA" Then TxtNomeFilial.Caption = "MARANHAO"
If TxtBuscaFilial.Text = "MG" Then TxtNomeFilial.Caption = "MINAS GERAIS"
If TxtBuscaFilial.Text = "MS" Then TxtNomeFilial.Caption = "MATO GROSSO DO SUL"
If TxtBuscaFilial.Text = "MT" Then TxtNomeFilial.Caption = "MATO GROSSO"
If TxtBuscaFilial.Text = "PA" Then TxtNomeFilial.Caption = "PARA"
If TxtBuscaFilial.Text = "PB" Then TxtNomeFilial.Caption = "PARAIBA"
If TxtBuscaFilial.Text = "PE" Then TxtNomeFilial.Caption = "PERNAMBUCO"
If TxtBuscaFilial.Text = "PI" Then TxtNomeFilial.Caption = "PIAUI"
If TxtBuscaFilial.Text = "PR" Then TxtNomeFilial.Caption = "PARANA"
If TxtBuscaFilial.Text = "RJ" Then TxtNomeFilial.Caption = "RIO DE JANEIRO"
If TxtBuscaFilial.Text = "RN" Then TxtNomeFilial.Caption = "RIO GRANDE DO NORTE"
If TxtBuscaFilial.Text = "RO" Then TxtNomeFilial.Caption = "RONDONIA"
If TxtBuscaFilial.Text = "RR" Then TxtNomeFilial.Caption = "RORAIMA"
If TxtBuscaFilial.Text = "RS" Then TxtNomeFilial.Caption = "RIO GRANDE DO SUL"
If TxtBuscaFilial.Text = "SC" Then TxtNomeFilial.Caption = "SANTA CATARINA"
If TxtBuscaFilial.Text = "SE" Then TxtNomeFilial.Caption = "SERGIPE"
If TxtBuscaFilial.Text = "SP" Then TxtNomeFilial.Caption = "SAO PAULO"
If TxtBuscaFilial.Text = "TO" Then TxtNomeFilial.Caption = "TOCANTINS"


'If Len(Trim(TxtBuscaFilial.Text)) > 0 Then
'    TxtBuscaFilial.Text = Trim(String(2 - Len(Trim(Str(Val(TxtBuscaFilial.Text)))), "0")) & Trim(Str(Val(TxtBuscaFilial.Text)))
'    If de_informa.rsSelFiliais.State = 1 Then de_informa.rsSelFiliais.Close
'    de_informa.SelFiliais TxtBuscaFilial.Text
'
'    If de_informa.rsSelFiliais.RecordCount > 0 Then
'        If IsNull(de_informa.rsSelFiliais.Fields("filial")) = False Then TxtFilial.Caption = de_informa.rsSelFiliais.Fields("filial")
'        If IsNull(de_informa.rsSelFiliais.Fields("nomefilial")) = False Then TxtNomeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("nomefilial"))
'        If IsNull(de_informa.rsSelFiliais.Fields("cgc")) = False Then TxtCGCFilial.Caption = de_informa.rsSelFiliais.Fields("cgc")
'        If IsNull(de_informa.rsSelFiliais.Fields("inscrest")) = False Then TxtInscrEstFilial.Caption = de_informa.rsSelFiliais.Fields("inscrest")
'        If IsNull(de_informa.rsSelFiliais.Fields("cidade")) = False Then TxtCidadeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("cidade"))
'        If IsNull(de_informa.rsSelFiliais.Fields("uf")) = False Then TxtUFFilial.Caption = de_informa.rsSelFiliais.Fields("uf")
'        If IsNull(de_informa.rsSelFiliais.Fields("licensaIATA")) = False Then TxtLicensaFilial.Caption = de_informa.rsSelFiliais.Fields("licensaIATA")
'        If IsNull(de_informa.rsSelFiliais.Fields("siglaIATA")) = False Then TxtSiglaFilial.Caption = de_informa.rsSelFiliais.Fields("siglaIATA")
'    DoEvents
'    End If
'End If
End Sub

Private Sub TxtBuscaSiglaCia_Change()
If Len(Trim(TxtBuscaSiglaCia.Text)) > 0 Then
TxtBuscaSiglaCia.Text = UCase(TxtBuscaSiglaCia.Text)
TxtBuscaSiglaCia.SelStart = Len(TxtBuscaSiglaCia.Text)
End If
End Sub

Private Sub TxtBuscaSiglaCia_GotFocus()
TxtBuscaSiglaCia.SelStart = 0
TxtBuscaSiglaCia.SelLength = 100
End Sub

Private Sub TxtBuscaSiglaCia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub TxtBuscaSiglaCia_LostFocus()
If Len(Trim(TxtBuscaSiglaCia.Text)) > 0 Then
    If de_informa.rsSelCiaAerea.State = 1 Then de_informa.rsSelCiaAerea.Close
    de_informa.SelCiaAerea Trim(UCase(TxtBuscaSiglaCia.Text))
    
        If de_informa.rsSelCiaAerea.RecordCount > 0 Then
        TxtSiglaCiaAerea.Text = UCase(de_informa.rsSelCiaAerea.Fields("codcia"))
        TxtNomeCiaAerea.Caption = PriMaiuscula(de_informa.rsSelCiaAerea.Fields("fantasia"))
        TxtCGCCiaAerea.Caption = de_informa.rsSelCiaAerea.Fields("cgc")
        TxtInscrEstCiaAerea.Caption = de_informa.rsSelCiaAerea.Fields("inscrest")
        Else
        TxtSiglaCiaAerea.Text = ""
        TxtNomeCiaAerea.Caption = ""
        TxtCGCCiaAerea.Caption = ""
        TxtInscrEstCiaAerea.Caption = ""
        MsgBox "Sigla de Cia. Aérea não encontrada! Por favor, tente novamente...", vbCritical, ""
        TxtBuscaSiglaCia.SetFocus
        End If
Else
TxtNomeCiaAerea.Caption = ""
TxtCGCCiaAerea.Caption = ""
TxtInscrEstCiaAerea.Caption = ""
End If
End Sub

Private Sub TxtNomeArquivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
