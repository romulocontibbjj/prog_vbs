VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManifesto 
   Caption         =   "Manifestos"
   ClientHeight    =   7230
   ClientLeft      =   1260
   ClientTop       =   930
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frmManifesto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Filial de Emissão do Manif."
      Height          =   615
      Left            =   60
      TabIndex        =   34
      Top             =   720
      Width           =   2115
      Begin VB.TextBox TxtEmissaoFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   6
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox TxtNomeFilial 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         MaxLength       =   50
         TabIndex        =   35
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Aeroporto de Origem"
      Height          =   615
      Left            =   60
      TabIndex        =   32
      Top             =   60
      Width           =   2715
      Begin VB.TextBox TxtSiglaOrigem 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtAeroporto 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   660
         MaxLength       =   50
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cia. Aérea"
      Height          =   615
      Left            =   2880
      TabIndex        =   30
      Top             =   0
      Width           =   2655
      Begin VB.TextBox TxtNomeCia 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         MaxLength       =   50
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtManualCia 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "AWBs a serem Manifestados"
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
      Height          =   4395
      Left            =   60
      TabIndex        =   26
      Top             =   2760
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid FlexAdVal1 
         Height          =   3735
         Left            =   60
         TabIndex        =   27
         Top             =   180
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid FlexAdVal1Total 
         Height          =   375
         Left            =   60
         TabIndex        =   28
         Top             =   3960
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "AD Valorem 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2415
      Left            =   60
      TabIndex        =   23
      Top             =   7380
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid FlexAdVal2Total 
         Height          =   375
         Left            =   60
         TabIndex        =   24
         Top             =   1980
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid FlexAdVal2 
         Height          =   1755
         Left            =   60
         TabIndex        =   25
         Top             =   180
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3096
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame3"
      Height          =   1275
      Left            =   60
      TabIndex        =   19
      Top             =   1440
      Width           =   7095
      Begin VB.TextBox TxtAjudante2 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   12
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox TxtAjudante1 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   11
         Top             =   540
         Width           =   5895
      End
      Begin VB.TextBox TxtMotorista 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   10
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2º Ajudante"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   885
         Width           =   825
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1º Ajudante"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motorista"
         Height          =   195
         Left            =   330
         TabIndex        =   20
         Top             =   285
         Width           =   645
      End
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   435
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Frame FraDadosManual 
      Caption         =   "Filial / AWB"
      Height          =   615
      Left            =   5640
      TabIndex        =   16
      Top             =   0
      Width           =   2655
      Begin VB.TextBox TxtDig 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         MaxLength       =   1
         TabIndex        =   4
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox TxtManualNumero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   600
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtManualFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   7260
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton CmdManifestar 
      Caption         =   "Manifestar"
      Height          =   495
      Left            =   7260
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Veículo"
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   720
      Width           =   7155
      Begin VB.TextBox TxtProprietario 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4140
         MaxLength       =   50
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox TxtCodigo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox TxtPlaca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   3240
         TabIndex        =   29
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Placa"
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         Top             =   285
         Width           =   405
      End
   End
   Begin VB.Menu mnuAdval2 
      Caption         =   "FlexAdVal2"
      Visible         =   0   'False
      Begin VB.Menu mnuDelAdVal2 
         Caption         =   "Deletar o AWB desta linha"
      End
   End
   Begin VB.Menu mnuAdval1 
      Caption         =   "FlexAdval1"
      Visible         =   0   'False
      Begin VB.Menu mnuDelAdVAl1 
         Caption         =   "Deletar o AWB desta linha"
      End
   End
End
Attribute VB_Name = "frmManifesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Aeroporto As String
Public CIA As String
Public Filial As String
Public Botao As Integer
Public xRow As Integer


Private Sub CmdBuscar_Click()
Dim xCodAwb As String
Dim X As Integer
Dim xRec As Double
Dim xVols As Double
Dim xPesoReal As Double
Dim xPesoTaxado As Double
Dim xFrete As Double

        If Len(Trim(TxtSiglaOrigem.Text)) = 0 Or Len(Trim(TxtManualFilial.Text)) = 0 Or Len(Trim(TxtManualNumero.Text)) = 0 Or Len(Trim(TxtManualCia.Text)) = 0 Then
        TxtEmissaoFilial.SetFocus
        Exit Sub
        End If
    
    'CodAwb = TxtFilial.Text & TxtSigla.Text & String(11 - Len(Trim(Str(Val(TxtAwb.Text)))), "0") & Str(Val(TxtAwb.Text))
    'xCodAwb = String(2 - Len(Trim(TxtManualFilial.Text)), "0") & Trim(TxtManualFilial.Text) & UCase(Trim(TxtManualCia.Text)) & String(11 - Len(Trim(Str(Val(TxtManualNumero.Text)))), "0") & Str(Val(TxtManualNumero.Text))
    xCodAwb = Trim(TxtManualFilial.Text) & UCase(Trim(TxtManualCia.Text)) & String(10 - Len(Trim(Str(Val(TxtManualNumero.Text)))), "0") & Trim(Str(Val(TxtManualNumero.Text))) & Trim(TxtDig.Text)
        
    If de_informa.rsConsultaAWBManifesto.State = 1 Then de_informa.rsConsultaAWBManifesto.Close
    de_informa.ConsultaAWBmanifesto xCodAwb, TxtSiglaOrigem.Text
        
        If de_informa.rsConsultaAWBManifesto.RecordCount = 0 Then
        MsgBox "O AWB informado para esta origem não foi encontrado. Verifique os dados do AWB e a origem que foi emitido.", vbExclamation, ""
        TxtManualFilial.Text = ""
        TxtManualNumero.Text = ""
        TxtDig.Text = ""
        TxtManualFilial.SetFocus
        Exit Sub
        End If
        
        For X = 1 To FlexAdVal1.Rows - 1
            If FlexAdVal1.TextMatrix(X, 2) = de_informa.rsConsultaAWBManifesto.Fields("awb") & "-" & de_informa.rsConsultaAWBManifesto.Fields("DIG") Then
            MsgBox "Você já incluiu este AWB.", vbExclamation, ""
            Exit Sub
            End If
        Next
        
        For X = 1 To FlexAdVal2.Rows - 1
            If FlexAdVal2.TextMatrix(X, 2) = de_informa.rsConsultaAWBManifesto.Fields("awb") Then
            MsgBox "Você já incluiu este AWB.", vbExclamation, ""
            Exit Sub
            End If
        Next
        
    
            FlexAdVal1.Rows = FlexAdVal1.Rows + 1
            X = FlexAdVal1.Rows - 2
            FlexAdVal1.TextMatrix(X, 0) = de_informa.rsConsultaAWBManifesto.Fields("filial")
            FlexAdVal1.TextMatrix(X, 1) = de_informa.rsConsultaAWBManifesto.Fields("cia")
            FlexAdVal1.TextMatrix(X, 2) = de_informa.rsConsultaAWBManifesto.Fields("awb") & "-" & de_informa.rsConsultaAWBManifesto.Fields("dig")
            FlexAdVal1.TextMatrix(X, 3) = de_informa.rsConsultaAWBManifesto.Fields("cidadedes")
            FlexAdVal1.TextMatrix(X, 4) = de_informa.rsConsultaAWBManifesto.Fields("siglades")
            FlexAdVal1.TextMatrix(X, 5) = de_informa.rsConsultaAWBManifesto.Fields("volumes")
            FlexAdVal1.TextMatrix(X, 6) = de_informa.rsConsultaAWBManifesto.Fields("pesoreal")
                If de_informa.rsConsultaAWBManifesto.Fields("pesoreal") > de_informa.rsConsultaAWBManifesto.Fields("pesocubado") Then
                FlexAdVal1.TextMatrix(X, 7) = de_informa.rsConsultaAWBManifesto.Fields("pesoreal")
                Else
                FlexAdVal1.TextMatrix(X, 7) = de_informa.rsConsultaAWBManifesto.Fields("pesocubado")
                End If
            FlexAdVal1.TextMatrix(X, 8) = de_informa.rsConsultaAWBManifesto.Fields("fretetotal")
            FlexAdVal1.TextMatrix(X, 9) = de_informa.rsConsultaAWBManifesto.Fields("perecivel")
            FlexAdVal1.TextMatrix(X, 10) = de_informa.rsConsultaAWBManifesto.Fields("tipoadval")
            
            xRec = 0
            xVols = 0
            xPesoReal = 0
            xPesoTaxado = 0
            xFrete = 0
                For X = 1 To FlexAdVal1.Rows - 1
                    If Len(Trim(FlexAdVal1.TextMatrix(X, 0))) > 0 Then
                    xRec = xRec + 1
                    xVols = xVols + FlexAdVal1.TextMatrix(X, 5)
                    xPesoReal = xPesoReal + FlexAdVal1.TextMatrix(X, 6)
                    xPesoTaxado = xPesoTaxado + FlexAdVal1.TextMatrix(X, 7)
                    xFrete = xFrete + FlexAdVal1.TextMatrix(X, 8)
                    End If
                Next
                
            FlexAdVal1Total.TextMatrix(0, 0) = xRec
            FlexAdVal1Total.TextMatrix(0, 1) = ""
            FlexAdVal1Total.TextMatrix(0, 2) = ""
            FlexAdVal1Total.TextMatrix(0, 3) = ""
            FlexAdVal1Total.TextMatrix(0, 4) = ""
            FlexAdVal1Total.TextMatrix(0, 5) = xVols
            FlexAdVal1Total.TextMatrix(0, 6) = xPesoReal
            FlexAdVal1Total.TextMatrix(0, 7) = xPesoTaxado
            FlexAdVal1Total.TextMatrix(0, 8) = xFrete
TxtManualFilial.Text = ""
TxtManualNumero.Text = ""
TxtDig.Text = ""
TxtManualFilial.SetFocus
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdManifestar_Click()
Dim xFilialManifesto As String
Dim Linha As Integer
Dim xCodAwb As String
Dim X As Integer
Dim Limite As Integer
Dim TotPag As Integer
Dim Pag As Integer
Dim xlinha As String


    If Val(FlexAdVal1Total.TextMatrix(0, 0)) = 0 Then
    MsgBox "Você não inseriu AWB algum para ser manifestado.", vbCritical, ""
    TxtManualFilial.SetFocus
    Exit Sub
    ElseIf Len(Trim(TxtPlaca.Text)) = 0 Then
    MsgBox "Você deve inserir a placa do veículo.", vbInformation, ""
    TxtPlaca.SetFocus
    Exit Sub
    ElseIf Len(Trim(TxtProprietario.Text)) = 0 Then
    MsgBox "Você deve inserir o nome do proprietário do veículo.", vbInformation, ""
    Exit Sub
    ElseIf Len(Trim(TxtMotorista.Text)) = 0 Then
    MsgBox "Você deve inserir o nome do motorista do veículo.", vbInformation, ""
    TxtMotorista.SetFocus
    Exit Sub
    ElseIf Len(Trim(TxtSiglaOrigem.Text)) = 0 Then
    MsgBox "Você deve inserir o aeroporto de Origem.", vbInformation, ""
    Exit Sub
    ElseIf Len(Trim(TxtManualCia.Text)) = 0 Then
    MsgBox "Você deve inserir a Cia. Aérea..", vbInformation, ""
    Exit Sub
    ElseIf Len(Trim(TxtNomeFilial.Text)) = 0 Then
    MsgBox "Você deve inserir a Filial.", vbInformation, ""
    Exit Sub
    End If
    
'CONFIGURACAO DE IMPRESSORAS - Inicio
Dim SETIMPLinha As String
Dim SETIMPLinhaPC As String
Dim SETIMPImpressoraAtual As Printer
Dim SETIMPAchouIMP As Boolean
Dim NomeMicro As String

    If Dir("c:\printer.cfg") = "" Then
    MsgBox "Você não possui o arquivo de configuração de impressoras. Antes de continuar, é imprescindível que você configure as configure.", vbExclamation, "IMPRESSORAS"
    frmControleImpressoras.Show 1
    End If

AchouIMP = False
    
    If Dir("c:\printer.cfg") <> "" Then
        Open "c:\printer.cfg" For Input As #1
        Do Until EOF(1)
            Line Input #1, SETIMPxLinha
            If Mid(SETIMPxLinha, 1, 3) = "MNF" Then
            SETIMPImpressoraPadrao = Mid(SETIMPxLinha, 5)
            SETIMPAchouIMP = True
            Exit Do
            End If
        Loop
        Close #1
    End If
    
    
    If SETIMPAchouIMP = False Then
    MsgBox "Não existe impressora configurada para esta operação. Corrija este problema indo ao menu Configurações e depois em Impressoras e configure em qual impressora os AWBs deverão ser impressos.", vbCritical, "ERRO!"
    Exit Sub
    End If
    
    For Each SETIMPImpressoraAtual In Printers
        If SETIMPImpressoraAtual.DeviceName = SETIMPImpressoraPadrao Then
            Set Printer = SETIMPImpressoraAtual
            DoEvents
            Exit For
        End If
    Next

    If Mid(SETIMPImpressoraPadrao, 1, 1) <> "\" Then
    SETIMPImpressoraPadrao = "LPT1"
    End If

'CONFIGURACAO DE IMPRESSORAS - Fim
    
frmManifesto.MousePointer = 11
DoEvents
de_informa.cn_informa.BeginTrans

Call AdValorem1NaoPerecivel
Call AdValorem1Perecivel
Call AdValorem2NaoPerecivel
Call AdValorem2Perecivel

de_informa.cn_informa.CommitTrans

Call LimpaTela(Me)
FlexAdVal1.Clear
FlexAdVal1Total.Clear
FlexAdVal2.Clear
FlexAdVal2Total.Clear

FlexAdVal1.Rows = 2
FlexAdVal1Total.Rows = 1
FlexAdVal1.Cols = 11
FlexAdVal1Total.Cols = 11
FlexAdVal1.FixedRows = 1
FlexAdVal1.FixedCols = 0
FlexAdVal1Total.FixedRows = 0
FlexAdVal1Total.FixedCols = 0
FlexAdVal1.ColWidth(0) = 500
FlexAdVal1.ColWidth(1) = 500
FlexAdVal1.ColWidth(2) = 1000
FlexAdVal1.ColWidth(3) = 2000
FlexAdVal1.ColWidth(4) = 800
FlexAdVal1.ColWidth(5) = 800
FlexAdVal1.ColWidth(6) = 1000
FlexAdVal1.ColWidth(7) = 1000
FlexAdVal1.ColWidth(8) = 1000
FlexAdVal1.ColWidth(9) = 200
FlexAdVal1.ColWidth(10) = 200
FlexAdVal1Total.ColWidth(0) = 500
FlexAdVal1Total.ColWidth(1) = 500
FlexAdVal1Total.ColWidth(2) = 1000
FlexAdVal1Total.ColWidth(3) = 2000
FlexAdVal1Total.ColWidth(4) = 800
FlexAdVal1Total.ColWidth(5) = 800
FlexAdVal1Total.ColWidth(6) = 1000
FlexAdVal1Total.ColWidth(7) = 1000
FlexAdVal1Total.ColWidth(8) = 1000
FlexAdVal1Total.ColWidth(9) = 200
FlexAdVal1Total.ColWidth(10) = 200
FlexAdVal1.TextMatrix(0, 0) = "Filial"
FlexAdVal1.TextMatrix(0, 1) = "Cia."
FlexAdVal1.TextMatrix(0, 2) = "AWB"
FlexAdVal1.TextMatrix(0, 3) = "Destino"
FlexAdVal1.TextMatrix(0, 4) = "Sigla"
FlexAdVal1.TextMatrix(0, 5) = "Vols."
FlexAdVal1.TextMatrix(0, 6) = "Peso Real"
FlexAdVal1.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal1.TextMatrix(0, 8) = "Frete"
FlexAdVal1.TextMatrix(0, 9) = "P"
FlexAdVal1.TextMatrix(0, 10) = "A"

FlexAdVal2.Rows = 2
FlexAdVal2Total.Rows = 1
FlexAdVal2.Cols = 11
FlexAdVal2Total.Cols = 11
FlexAdVal2.FixedRows = 1
FlexAdVal2.FixedCols = 0
FlexAdVal2Total.FixedRows = 0
FlexAdVal2Total.FixedCols = 0
FlexAdVal2.ColWidth(0) = 500
FlexAdVal2.ColWidth(1) = 500
FlexAdVal2.ColWidth(2) = 1000
FlexAdVal2.ColWidth(3) = 2000
FlexAdVal2.ColWidth(4) = 800
FlexAdVal2.ColWidth(5) = 800
FlexAdVal2.ColWidth(6) = 1000
FlexAdVal2.ColWidth(7) = 1000
FlexAdVal2.ColWidth(8) = 1000
FlexAdVal2.ColWidth(9) = 200
FlexAdVal2.ColWidth(10) = 200
FlexAdVal2Total.ColWidth(0) = 500
FlexAdVal2Total.ColWidth(1) = 500
FlexAdVal2Total.ColWidth(2) = 1000
FlexAdVal2Total.ColWidth(3) = 2000
FlexAdVal2Total.ColWidth(4) = 800
FlexAdVal2Total.ColWidth(5) = 800
FlexAdVal2Total.ColWidth(6) = 1000
FlexAdVal2Total.ColWidth(7) = 1000
FlexAdVal2Total.ColWidth(8) = 1000
FlexAdVal2Total.ColWidth(9) = 200
FlexAdVal2Total.ColWidth(10) = 200
FlexAdVal2.TextMatrix(0, 0) = "Filial"
FlexAdVal2.TextMatrix(0, 1) = "Cia."
FlexAdVal2.TextMatrix(0, 2) = "AWB"
FlexAdVal2.TextMatrix(0, 3) = "Destino"
FlexAdVal2.TextMatrix(0, 4) = "Sigla"
FlexAdVal2.TextMatrix(0, 5) = "Vols."
FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal2.TextMatrix(0, 8) = "Frete"
FlexAdVal2.TextMatrix(0, 9) = "P"
FlexAdVal2.TextMatrix(0, 10) = "A"

frmManifesto.MousePointer = 0
DoEvents
End Sub

Private Sub FlexAdVal1_Click()
xRow = FlexAdVal1.Row
    If Val(FlexAdVal1Total.TextMatrix(0, 0)) > 0 Then
        If xRow > 0 Then
            If Botao = 2 Then
            PopupMenu mnuAdval1
            End If
        End If
    End If
End Sub

Private Sub FlexAdVal1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Botao = Button
End Sub

Private Sub Form_Load()
'frmManifesto.StartUpPosition = 2

FlexAdVal1.Clear
FlexAdVal1Total.Clear
FlexAdVal2.Clear
FlexAdVal2Total.Clear

FlexAdVal1.Rows = 2
FlexAdVal1Total.Rows = 1
FlexAdVal1.Cols = 11
FlexAdVal1Total.Cols = 11
FlexAdVal1.FixedRows = 1
FlexAdVal1.FixedCols = 0
FlexAdVal1Total.FixedRows = 0
FlexAdVal1Total.FixedCols = 0
FlexAdVal1.ColWidth(0) = 500
FlexAdVal1.ColWidth(1) = 500
FlexAdVal1.ColWidth(2) = 1000
FlexAdVal1.ColWidth(3) = 2000
FlexAdVal1.ColWidth(4) = 800
FlexAdVal1.ColWidth(5) = 800
FlexAdVal1.ColWidth(6) = 1000
FlexAdVal1.ColWidth(7) = 1000
FlexAdVal1.ColWidth(8) = 1000
FlexAdVal1.ColWidth(9) = 200
FlexAdVal1.ColWidth(10) = 200
FlexAdVal1Total.ColWidth(0) = 500
FlexAdVal1Total.ColWidth(1) = 500
FlexAdVal1Total.ColWidth(2) = 1000
FlexAdVal1Total.ColWidth(3) = 2000
FlexAdVal1Total.ColWidth(4) = 800
FlexAdVal1Total.ColWidth(5) = 800
FlexAdVal1Total.ColWidth(6) = 1000
FlexAdVal1Total.ColWidth(7) = 1000
FlexAdVal1Total.ColWidth(8) = 1000
FlexAdVal1Total.ColWidth(9) = 200
FlexAdVal1Total.ColWidth(10) = 200
FlexAdVal1.TextMatrix(0, 0) = "Filial"
FlexAdVal1.TextMatrix(0, 1) = "Cia."
FlexAdVal1.TextMatrix(0, 2) = "AWB"
FlexAdVal1.TextMatrix(0, 3) = "Destino"
FlexAdVal1.TextMatrix(0, 4) = "Sigla"
FlexAdVal1.TextMatrix(0, 5) = "Vols."
FlexAdVal1.TextMatrix(0, 6) = "Peso Real"
FlexAdVal1.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal1.TextMatrix(0, 8) = "Frete"
FlexAdVal1.TextMatrix(0, 9) = "P"
FlexAdVal1.TextMatrix(0, 10) = "A"

FlexAdVal2.Rows = 2
FlexAdVal2Total.Rows = 1
FlexAdVal2.Cols = 11
FlexAdVal2Total.Cols = 11
FlexAdVal2.FixedRows = 1
FlexAdVal2.FixedCols = 0
FlexAdVal2Total.FixedRows = 0
FlexAdVal2Total.FixedCols = 0
FlexAdVal2.ColWidth(0) = 500
FlexAdVal2.ColWidth(1) = 500
FlexAdVal2.ColWidth(2) = 1000
FlexAdVal2.ColWidth(3) = 2000
FlexAdVal2.ColWidth(4) = 800
FlexAdVal2.ColWidth(5) = 800
FlexAdVal2.ColWidth(6) = 1000
FlexAdVal2.ColWidth(7) = 1000
FlexAdVal2.ColWidth(8) = 1000
FlexAdVal2.ColWidth(9) = 200
FlexAdVal2.ColWidth(10) = 200
FlexAdVal2Total.ColWidth(0) = 500
FlexAdVal2Total.ColWidth(1) = 500
FlexAdVal2Total.ColWidth(2) = 1000
FlexAdVal2Total.ColWidth(3) = 2000
FlexAdVal2Total.ColWidth(4) = 800
FlexAdVal2Total.ColWidth(5) = 800
FlexAdVal2Total.ColWidth(6) = 1000
FlexAdVal2Total.ColWidth(7) = 1000
FlexAdVal2Total.ColWidth(8) = 1000
FlexAdVal2Total.ColWidth(9) = 200
FlexAdVal2Total.ColWidth(10) = 200
FlexAdVal2.TextMatrix(0, 0) = "Filial"
FlexAdVal2.TextMatrix(0, 1) = "Cia."
FlexAdVal2.TextMatrix(0, 2) = "AWB"
FlexAdVal2.TextMatrix(0, 3) = "Destino"
FlexAdVal2.TextMatrix(0, 4) = "Sigla"
FlexAdVal2.TextMatrix(0, 5) = "Vols."
FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal2.TextMatrix(0, 8) = "Frete"
FlexAdVal2.TextMatrix(0, 9) = "P"
FlexAdVal2.TextMatrix(0, 10) = "A"
End Sub

Private Sub OptAuto_Click()
    If OptAuto.Value = True Then
    TxtManualFilial.Text = ""
    TxtManualNumero.Text = ""
    TxtManualCia.Text = ""
    FraDadosManual.Visible = False
    TxtAutoFilial.Text = ""
    TxtAutoAeroporto.Text = ""
    TxtAutoCia.Text = ""
    FraDadosAuto.Visible = True
    Else
    TxtManualFilial.Text = ""
    TxtManualNumero.Text = ""
    TxtManualCia.Text = ""
    FraDadosManual.Visible = True
    TxtAutoFilial.Text = ""
    TxtAutoAeroporto.Text = ""
    TxtAutoCia.Text = ""
    FraDadosAuto.Visible = False
    End If
End Sub

Private Sub OptManual_Click()
    If OptAuto.Value = True Then
    TxtManualFilial.Text = ""
    TxtManualNumero.Text = ""
    TxtManualCia.Text = ""
    FraDadosManual.Visible = False
    TxtAutoFilial.Text = ""
    TxtAutoAeroporto.Text = ""
    TxtAutoCia.Text = ""
    FraDadosAuto.Visible = True
    Else
    TxtManualFilial.Text = ""
    TxtManualNumero.Text = ""
    TxtManualCia.Text = ""
    FraDadosManual.Visible = True
    TxtAutoFilial.Text = ""
    TxtAutoAeroporto.Text = ""
    TxtAutoCia.Text = ""
    FraDadosAuto.Visible = False
    End If
End Sub

Private Sub TxtAutoFilial_LostFocus()
TxtAutoFilial.Text = String(2 - Len(TxtAutoFilial.Text), "0") & TxtAutoFilial.Text
End Sub

Private Sub mnuDELAdval1_Click()
    If Val(FlexAdVal1.TextMatrix(xRow, 0)) > 0 Then
    FlexAdVal1.RemoveItem (xRow)
    xRec = 0
    xVols = 0
    xPesoReal = 0
    xPesoTaxado = 0
    xFrete = 0
        For X = 1 To FlexAdVal1.Rows - 1
            If Len(Trim(FlexAdVal1.TextMatrix(X, 0))) > 0 Then
            xRec = xRec + 1
            xVols = xVols + FlexAdVal1.TextMatrix(X, 5)
            xPesoReal = xPesoReal + FlexAdVal1.TextMatrix(X, 6)
            xPesoTaxado = xPesoTaxado + FlexAdVal1.TextMatrix(X, 7)
            xFrete = xFrete + FlexAdVal1.TextMatrix(X, 8)
            End If
        Next
        
    FlexAdVal1Total.TextMatrix(0, 0) = xRec
    FlexAdVal1Total.TextMatrix(0, 1) = ""
    FlexAdVal1Total.TextMatrix(0, 2) = ""
    FlexAdVal1Total.TextMatrix(0, 3) = ""
    FlexAdVal1Total.TextMatrix(0, 4) = ""
    FlexAdVal1Total.TextMatrix(0, 5) = xVols
    FlexAdVal1Total.TextMatrix(0, 6) = xPesoReal
    FlexAdVal1Total.TextMatrix(0, 7) = xPesoTaxado
    FlexAdVal1Total.TextMatrix(0, 8) = xFrete
    End If
End Sub

Private Sub TxtAjudante1_Change()
TxtAjudante1.Text = UCase(TxtAjudante1.Text)
TxtAjudante1.SelStart = Len(TxtAjudante1.Text)
End Sub

Private Sub TxtAjudante1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If

End Sub

Private Sub TxtAjudante2_Change()
TxtAjudante2.Text = UCase(TxtAjudante2.Text)
TxtAjudante2.SelStart = Len(TxtAjudante2.Text)
End Sub

Private Sub TxtAjudante2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub TxtCodigo_LostFocus()
If Len(Trim(TxtCodigo.Text)) > 0 Then
TxtPlaca.Text = ""
    If de_informa.rsVeiculoCOD.State = 1 Then de_informa.rsVeiculoCOD.Close
    de_informa.Veiculocod TxtCodigo.Text
    
    If de_informa.rsVeiculoCOD.RecordCount > 0 Then
    TxtProprietario.Enabled = False
    TxtProprietario.BackColor = xBranco
    TxtProprietario.Text = ""
    TxtPlaca.Text = de_informa.rsVeiculoCOD.Fields("placa")
    TxtProprietario.Text = de_informa.rsVeiculoCOD("proprietario")
    Else
    MsgBox "Código de Veículo não encontrado! Procure pela placa do mesmo.", vbExclamation, ""
    End If
End If
End Sub

Private Sub TxtDig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If
End Sub

Private Sub TxtEmissaoFilial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtEmissaoFilial_LostFocus()
If Len(Trim(TxtEmissaoFilial.Text)) > 0 Then
TxtEmissaoFilial.Text = String(2 - Len(TxtEmissaoFilial.Text), "0") & TxtEmissaoFilial.Text
    If de_informa.rsSelFiliais.State = 1 Then de_informa.rsSelFiliais.Close
    de_informa.SelFiliais TxtEmissaoFilial.Text
    
    TxtEmissaoFilial.Text = ""
    TxtNomeFilial.Text = ""
    If de_informa.rsSelFiliais.RecordCount = 0 Then
    MsgBox "Filial não encontrada!", vbCritical, ""
    TxtEmissaoFilial.SetFocus
    Exit Sub
    Else
    TxtNomeFilial.Text = de_informa.rsSelFiliais.Fields("filial") & " - " & PriMaiuscula(de_informa.rsSelFiliais.Fields("nomefilial"))
    End If
End If
End Sub

Private Sub TxtManualCia_Change()
TxtManualCia.Text = UCase(TxtManualCia.Text)
TxtManualCia.SelStart = Len(TxtManualCia.Text)
End Sub

Private Sub TxtManualCia_GotFocus()
TxtManualCia.SelStart = 0
TxtManualCia.SelLength = 3
CIA = TxtManualCia.Text
End Sub

Private Sub TxtManualCia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If
End Sub

Private Sub TxtManualCia_LostFocus()
If TxtManualCia.Text <> CIA Then
    If Val(FlexAdVal1Total.TextMatrix(0, 0)) > 0 Or Val(FlexAdVal2Total.TextMatrix(0, 0)) > 0 Then
        If MsgBox("Você está informando uma Cia. Aérea diferente das dos AWB inseridos. Se você escolher SIM, toda a sua seleção de AWBs será apagada. Deseja fazer isto?", vbYesNo + vbExclamation, "ATENÇÃO") = vbYes Then
        FlexAdVal1.Clear
        FlexAdVal1Total.Clear
        FlexAdVal2.Clear
        FlexAdVal2Total.Clear
        
        FlexAdVal1.Rows = 2
        FlexAdVal1Total.Rows = 1
        FlexAdVal1.Cols = 9
        FlexAdVal1Total.Cols = 9
        FlexAdVal1.FixedRows = 1
        FlexAdVal1.FixedCols = 0
        FlexAdVal1Total.FixedRows = 0
        FlexAdVal1Total.FixedCols = 0
        FlexAdVal1.ColWidth(0) = 500
        FlexAdVal1.ColWidth(1) = 500
        FlexAdVal1.ColWidth(2) = 1000
        FlexAdVal1.ColWidth(3) = 2000
        FlexAdVal1.ColWidth(4) = 1000
        FlexAdVal1.ColWidth(5) = 1000
        FlexAdVal1.ColWidth(6) = 1000
        FlexAdVal1.ColWidth(7) = 1000
        FlexAdVal1.ColWidth(8) = 1000
        FlexAdVal1Total.ColWidth(0) = 500
        FlexAdVal1Total.ColWidth(1) = 500
        FlexAdVal1Total.ColWidth(2) = 1000
        FlexAdVal1Total.ColWidth(3) = 2000
        FlexAdVal1Total.ColWidth(4) = 1000
        FlexAdVal1Total.ColWidth(5) = 1000
        FlexAdVal1Total.ColWidth(6) = 1000
        FlexAdVal1Total.ColWidth(7) = 1000
        FlexAdVal1Total.ColWidth(8) = 1000
        FlexAdVal1.TextMatrix(0, 0) = "Filial"
        FlexAdVal1.TextMatrix(0, 1) = "Cia."
        FlexAdVal1.TextMatrix(0, 2) = "AWB"
        FlexAdVal1.TextMatrix(0, 3) = "Destino"
        FlexAdVal1.TextMatrix(0, 4) = "Sigla"
        FlexAdVal1.TextMatrix(0, 5) = "Vols."
        FlexAdVal1.TextMatrix(0, 6) = "Peso Real"
        FlexAdVal1.TextMatrix(0, 7) = "Peso Tax."
        FlexAdVal1.TextMatrix(0, 8) = "Frete"
        
        FlexAdVal2.Rows = 2
        FlexAdVal2Total.Rows = 1
        FlexAdVal2.Cols = 9
        FlexAdVal2Total.Cols = 9
        FlexAdVal2.FixedRows = 1
        FlexAdVal2.FixedCols = 0
        FlexAdVal2Total.FixedRows = 0
        FlexAdVal2Total.FixedCols = 0
        FlexAdVal2.ColWidth(0) = 500
        FlexAdVal2.ColWidth(1) = 500
        FlexAdVal2.ColWidth(2) = 1000
        FlexAdVal2.ColWidth(3) = 2000
        FlexAdVal2.ColWidth(4) = 1000
        FlexAdVal2.ColWidth(5) = 1000
        FlexAdVal2.ColWidth(6) = 1000
        FlexAdVal2.ColWidth(7) = 1000
        FlexAdVal2.ColWidth(8) = 1000
        FlexAdVal2Total.ColWidth(0) = 500
        FlexAdVal2Total.ColWidth(1) = 500
        FlexAdVal2Total.ColWidth(2) = 1000
        FlexAdVal2Total.ColWidth(3) = 2000
        FlexAdVal2Total.ColWidth(4) = 1000
        FlexAdVal2Total.ColWidth(5) = 1000
        FlexAdVal2Total.ColWidth(6) = 1000
        FlexAdVal2Total.ColWidth(7) = 1000
        FlexAdVal2Total.ColWidth(8) = 1000
        FlexAdVal2.TextMatrix(0, 0) = "Filial"
        FlexAdVal2.TextMatrix(0, 1) = "Cia."
        FlexAdVal2.TextMatrix(0, 2) = "AWB"
        FlexAdVal2.TextMatrix(0, 3) = "Destino"
        FlexAdVal2.TextMatrix(0, 4) = "Sigla"
        FlexAdVal2.TextMatrix(0, 5) = "Vols."
        FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
        FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
        FlexAdVal2.TextMatrix(0, 8) = "Frete"
        Else
        TxtManualCia.Text = CIA
        End If
    End If
    
    If de_informa.rsSelCiaAerea.State = 1 Then de_informa.rsSelCiaAerea.Close
    de_informa.SelCiaAerea TxtManualCia.Text
    TxtManualCia.Text = ""
        If de_informa.rsSelCiaAerea.RecordCount = 0 Then
        MsgBox "Companhia Aérea não encontrada!", vbCritical, ""
        Else
        TxtManualCia.Text = de_informa.rsSelCiaAerea.Fields("codcia")
        TxtNomeCia.Text = PriMaiuscula(de_informa.rsSelCiaAerea.Fields("fantasia"))
        End If
End If
End Sub


Private Sub TxtManualFilial_GotFocus()
TxtManualFilial.SelStart = 0
TxtManualFilial.SelLength = 10
Filial = TxtManualFilial.Text
End Sub

Private Sub TxtManualFilial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If
End Sub

Private Sub TxtManualFilial_LostFocus()
If Len(Trim(TxtManualFilial.Text)) > 0 Then
TxtManualFilial.Text = String(2 - Len(TxtManualFilial.Text), "0") & TxtManualFilial.Text
End If
End Sub

Private Sub TxtManualNumero_GotFocus()
TxtManualNumero.SelStart = 0
TxtManualNumero.SelLength = 15
End Sub

Private Sub TxtManualNumero_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub TxtManualNumero_LostFocus()

If Len(Trim(TxtManualFilial.Text)) > 0 And Len(Trim(TxtManualCia.Text)) > 0 And Len(Trim(TxtManualNumero.Text)) > 0 Then

If de_informa.rsConfereNumeroAWB.State = 1 Then de_informa.rsConfereNumeroAWB.Close
de_informa.ConfereNumeroAWB TxtManualCia.Text, TxtManualFilial.Text, TxtManualNumero.Text

    If de_informa.rsConfereNumeroAWB.RecordCount = 0 Then
    MsgBox "Este formulário não está cadastrado!.", vbCritical, ""
    TxtManualNumero.SetFocus
    Exit Sub
    ElseIf de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "C" Then
    MsgBox "O formulário para este AWB está cancelado. Para utilizá-lo, vá até o cadastro de formulários e descancele-o.", vbCritical, ""
    TxtManualNumero.SetFocus
    Exit Sub
    Else
    TxtDig.Text = de_informa.rsConfereNumeroAWB.Fields("dig")
    End If
Else
TxtManualNumero.Text = ""
TxtDig.Text = ""
End If

End Sub

Private Sub TxtMotorista_Change()
TxtMotorista.Text = UCase(TxtMotorista.Text)
TxtMotorista.SelStart = Len(TxtMotorista.Text)
End Sub

Private Sub TxtMotorista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If

End Sub

Private Sub TxtMotorista_LostFocus()
If Len(Trim(TxtMotorista.Text)) > 0 Then
If de_informa.rsMotorista.State = 1 Then de_informa.rsMotorista.Close
de_informa.Motorista TxtMotorista.Text & "%"

    If de_informa.rsMotorista.RecordCount > 0 Then
    frmManifestoFiltraMotorista.Show 1
    ElseIf de_informa.rsMotorista.RecordCount = 1 Then
    TxtMotorista.Text = de_informa.rsMotorista.Fields("nome")
    End If
End If
End Sub

Private Sub TxtPlaca_Change()
TxtPlaca.Text = UCase(TxtPlaca.Text)
TxtPlaca.SelStart = Len(TxtPlaca.Text)
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
    ElseIf KeyAscii = 45 Or KeyAscii = 32 Then
    KeyAscii = 0
    End If
End Sub

Private Sub TxtPlaca_LostFocus()
If Len(Trim(TxtPlaca.Text)) > 0 Then
TxtCodigo.Text = ""
    If de_informa.rsVeiculoPLACA.State = 1 Then de_informa.rsVeiculoPLACA.Close
    de_informa.Veiculoplaca TxtPlaca.Text
    
    If de_informa.rsVeiculoPLACA.RecordCount > 0 Then
    TxtProprietario.Enabled = False
    TxtProprietario.BackColor = xBranco
    TxtProprietario.Text = ""
    TxtCodigo.Text = de_informa.rsVeiculoPLACA.Fields("codigo")
    TxtPlaca.Text = de_informa.rsVeiculoPLACA.Fields("placa")
    TxtProprietario.Text = de_informa.rsVeiculoPLACA("proprietario")
    Else
    MsgBox "Placa de Veículo não encontrado! Insira manualmente o proprietário do veículo.", vbExclamation, ""
    TxtProprietario.Enabled = True
    TxtProprietario.BackColor = xAmarelo
    TxtProprietario.Text = ""
    TxtProprietario.SetFocus
    End If
End If
End Sub

Private Sub TxtProprietario_Change()
TxtProprietario.Text = UCase(TxtProprietario.Text)
TxtProprietario.SelStart = Len(TxtProprietario.Text)
End Sub

Private Sub TxtProprietario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{TAB}"
End If

End Sub

Private Sub TxtSiglaOrigem_Change()
TxtSiglaOrigem.Text = UCase(TxtSiglaOrigem.Text)
TxtSiglaOrigem.SelStart = Len(TxtSiglaOrigem.Text)
End Sub

Private Sub TxtSiglaOrigem_GotFocus()
TxtSiglaOrigem.SelStart = 0
TxtSiglaOrigem.SelLength = 10
Aeroporto = TxtSiglaOrigem.Text
End Sub

Private Sub TxtSiglaOrigem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtSiglaOrigem_LostFocus()
If TxtSiglaOrigem.Text <> Aeroporto Then
    If Val(FlexAdVal1Total.TextMatrix(0, 0)) > 0 Or Val(FlexAdVal2Total.TextMatrix(0, 0)) > 0 Then
        If MsgBox("Você está informando uma origem diferente das dos AWB inseridos. Se você escolher SIM, toda a sua seleção de AWBs será apagada. Deseja fazer isto?", vbYesNo + vbExclamation, "ATENÇÃO") = vbYes Then
        FlexAdVal1.Clear
        FlexAdVal1Total.Clear
        FlexAdVal2.Clear
        FlexAdVal2Total.Clear
        
        FlexAdVal1.Rows = 2
        FlexAdVal1Total.Rows = 1
        FlexAdVal1.Cols = 9
        FlexAdVal1Total.Cols = 9
        FlexAdVal1.FixedRows = 1
        FlexAdVal1.FixedCols = 0
        FlexAdVal1Total.FixedRows = 0
        FlexAdVal1Total.FixedCols = 0
        FlexAdVal1.ColWidth(0) = 500
        FlexAdVal1.ColWidth(1) = 500
        FlexAdVal1.ColWidth(2) = 1000
        FlexAdVal1.ColWidth(3) = 2000
        FlexAdVal1.ColWidth(4) = 1000
        FlexAdVal1.ColWidth(5) = 1000
        FlexAdVal1.ColWidth(6) = 1000
        FlexAdVal1.ColWidth(7) = 1000
        FlexAdVal1.ColWidth(8) = 1000
        FlexAdVal1Total.ColWidth(0) = 500
        FlexAdVal1Total.ColWidth(1) = 500
        FlexAdVal1Total.ColWidth(2) = 1000
        FlexAdVal1Total.ColWidth(3) = 2000
        FlexAdVal1Total.ColWidth(4) = 1000
        FlexAdVal1Total.ColWidth(5) = 1000
        FlexAdVal1Total.ColWidth(6) = 1000
        FlexAdVal1Total.ColWidth(7) = 1000
        FlexAdVal1Total.ColWidth(8) = 1000
        FlexAdVal1.TextMatrix(0, 0) = "Filial"
        FlexAdVal1.TextMatrix(0, 1) = "Cia."
        FlexAdVal1.TextMatrix(0, 2) = "AWB"
        FlexAdVal1.TextMatrix(0, 3) = "Destino"
        FlexAdVal1.TextMatrix(0, 4) = "Sigla"
        FlexAdVal1.TextMatrix(0, 5) = "Vols."
        FlexAdVal1.TextMatrix(0, 6) = "Peso Real"
        FlexAdVal1.TextMatrix(0, 7) = "Peso Tax."
        FlexAdVal1.TextMatrix(0, 8) = "Frete"
        
        FlexAdVal2.Rows = 2
        FlexAdVal2Total.Rows = 1
        FlexAdVal2.Cols = 9
        FlexAdVal2Total.Cols = 9
        FlexAdVal2.FixedRows = 1
        FlexAdVal2.FixedCols = 0
        FlexAdVal2Total.FixedRows = 0
        FlexAdVal2Total.FixedCols = 0
        FlexAdVal2.ColWidth(0) = 500
        FlexAdVal2.ColWidth(1) = 500
        FlexAdVal2.ColWidth(2) = 1000
        FlexAdVal2.ColWidth(3) = 2000
        FlexAdVal2.ColWidth(4) = 1000
        FlexAdVal2.ColWidth(5) = 1000
        FlexAdVal2.ColWidth(6) = 1000
        FlexAdVal2.ColWidth(7) = 1000
        FlexAdVal2.ColWidth(8) = 1000
        FlexAdVal2Total.ColWidth(0) = 500
        FlexAdVal2Total.ColWidth(1) = 500
        FlexAdVal2Total.ColWidth(2) = 1000
        FlexAdVal2Total.ColWidth(3) = 2000
        FlexAdVal2Total.ColWidth(4) = 1000
        FlexAdVal2Total.ColWidth(5) = 1000
        FlexAdVal2Total.ColWidth(6) = 1000
        FlexAdVal2Total.ColWidth(7) = 1000
        FlexAdVal2Total.ColWidth(8) = 1000
        FlexAdVal2.TextMatrix(0, 0) = "Filial"
        FlexAdVal2.TextMatrix(0, 1) = "Cia."
        FlexAdVal2.TextMatrix(0, 2) = "AWB"
        FlexAdVal2.TextMatrix(0, 3) = "Destino"
        FlexAdVal2.TextMatrix(0, 4) = "Sigla"
        FlexAdVal2.TextMatrix(0, 5) = "Vols."
        FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
        FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
        FlexAdVal2.TextMatrix(0, 8) = "Frete"
        Else
        TxtSiglaOrigem.Text = Aeroporto
        End If
        
    End If
    
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    de_informa.SelAeroportoSigla TxtSiglaOrigem.Text
    
    TxtSiglaOrigem.Text = ""
        If de_informa.rsSelAeroportoSigla.RecordCount = 0 Then
        MsgBox "Aeroporto não encontrado!", vbCritical, ""
        Else
        TxtSiglaOrigem.Text = de_informa.rsSelAeroportoSigla.Fields("sigla")
        TxtAeroporto.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("localidade"))
        End If
    
End If
End Sub



Sub AdValorem1Perecivel()

'ADVALOREM 1 PERECIVEL
FlexAdVal2.Clear
FlexAdVal2Total.Clear

FlexAdVal2.Rows = 2
FlexAdVal2Total.Rows = 1
FlexAdVal2.Cols = 11
FlexAdVal2Total.Cols = 11
FlexAdVal2.FixedRows = 1
FlexAdVal2.FixedCols = 0
FlexAdVal2Total.FixedRows = 0
FlexAdVal2Total.FixedCols = 0
FlexAdVal2.ColWidth(0) = 500
FlexAdVal2.ColWidth(1) = 500
FlexAdVal2.ColWidth(2) = 1000
FlexAdVal2.ColWidth(3) = 2000
FlexAdVal2.ColWidth(4) = 800
FlexAdVal2.ColWidth(5) = 800
FlexAdVal2.ColWidth(6) = 1000
FlexAdVal2.ColWidth(7) = 1000
FlexAdVal2.ColWidth(8) = 1000
FlexAdVal2.ColWidth(9) = 200
FlexAdVal2.ColWidth(10) = 200
FlexAdVal2Total.ColWidth(0) = 500
FlexAdVal2Total.ColWidth(1) = 500
FlexAdVal2Total.ColWidth(2) = 1000
FlexAdVal2Total.ColWidth(3) = 2000
FlexAdVal2Total.ColWidth(4) = 800
FlexAdVal2Total.ColWidth(5) = 800
FlexAdVal2Total.ColWidth(6) = 1000
FlexAdVal2Total.ColWidth(7) = 1000
FlexAdVal2Total.ColWidth(8) = 1000
FlexAdVal2Total.ColWidth(9) = 200
FlexAdVal2Total.ColWidth(10) = 200
FlexAdVal2.TextMatrix(0, 0) = "Filial"
FlexAdVal2.TextMatrix(0, 1) = "Cia."
FlexAdVal2.TextMatrix(0, 2) = "AWB"
FlexAdVal2.TextMatrix(0, 3) = "Destino"
FlexAdVal2.TextMatrix(0, 4) = "Sigla"
FlexAdVal2.TextMatrix(0, 5) = "Vols."
FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal2.TextMatrix(0, 8) = "Frete"
FlexAdVal2.TextMatrix(0, 9) = "P"
FlexAdVal2.TextMatrix(0, 10) = "A"

    For X = 1 To FlexAdVal1.Rows - 1
        If FlexAdVal1.TextMatrix(X, 9) = "S" And Val(FlexAdVal1.TextMatrix(X, 10)) = 1 Then
        FlexAdVal2.Rows = FlexAdVal2.Rows + 1
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 0) = FlexAdVal1.TextMatrix(X, 0)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 1) = FlexAdVal1.TextMatrix(X, 1)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 2) = FlexAdVal1.TextMatrix(X, 2)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 3) = FlexAdVal1.TextMatrix(X, 3)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 4) = FlexAdVal1.TextMatrix(X, 4)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 5) = FlexAdVal1.TextMatrix(X, 5)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 6) = FlexAdVal1.TextMatrix(X, 6)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 7) = FlexAdVal1.TextMatrix(X, 7)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 8) = FlexAdVal1.TextMatrix(X, 8)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 9) = FlexAdVal1.TextMatrix(X, 9)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 10) = FlexAdVal1.TextMatrix(X, 10)
        End If
    Next
    
    xRec = 0
    xVols = 0
    xPesoReal = 0
    xPesoTaxado = 0
    xFrete = 0
        For X = 1 To FlexAdVal2.Rows - 1
            If Len(Trim(FlexAdVal2.TextMatrix(X, 0))) > 0 Then
            xRec = xRec + 1
            xVols = xVols + FlexAdVal2.TextMatrix(X, 5)
            xPesoReal = xPesoReal + FlexAdVal2.TextMatrix(X, 6)
            xPesoTaxado = xPesoTaxado + FlexAdVal2.TextMatrix(X, 7)
            xFrete = xFrete + FlexAdVal2.TextMatrix(X, 8)
            End If
        Next
    FlexAdVal2Total.TextMatrix(0, 0) = xRec
    FlexAdVal2Total.TextMatrix(0, 1) = ""
    FlexAdVal2Total.TextMatrix(0, 2) = ""
    FlexAdVal2Total.TextMatrix(0, 3) = ""
    FlexAdVal2Total.TextMatrix(0, 4) = ""
    FlexAdVal2Total.TextMatrix(0, 5) = xVols
    FlexAdVal2Total.TextMatrix(0, 6) = xPesoReal
    FlexAdVal2Total.TextMatrix(0, 7) = xPesoTaxado
    FlexAdVal2Total.TextMatrix(0, 8) = xFrete

If Val(FlexAdVal2Total.TextMatrix(0, 0)) > 0 Then
If de_informa.rsCapturaManifesto.State = 1 Then de_informa.rsCapturaManifesto.Close
de_informa.CapturaManifesto Mid(TxtNomeFilial.Text, 1, 2)

    If de_informa.rsCapturaManifesto.RecordCount > 0 Then
        If Not IsNull(de_informa.rsCapturaManifesto.Fields("manifesto")) Then
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))), "0") & Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))
        Else
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
        End If
    Else
    xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
    End If
    
    For Linha = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        If Val(FlexAdVal2.TextMatrix(Linha, 0)) > 0 Then
        xCodAwb = FlexAdVal2.TextMatrix(Linha, 0) & FlexAdVal2.TextMatrix(Linha, 1) & String(10 - Len(Trim(Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2))), "0") & Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2) & Mid(FlexAdVal2.TextMatrix(Linha, 2), Len(FlexAdVal2.TextMatrix(Linha, 2)))
        de_informa.InsAirManifesto xFilialManifesto, FlexAdVal2.TextMatrix(Linha, 0), Trim(Str(Val(Mid(xFilialManifesto, 3)))), DataHora("DATA"), _
        DataHora("HORA"), FlexAdVal2.TextMatrix(Linha, 10), FlexAdVal2.TextMatrix(Linha, 9), FlexAdVal2.TextMatrix(Linha, 1), UCase(TxtNomeCia.Text), xCodAwb, UCase(TxtAeroporto.Text), TxtSiglaOrigem.Text, _
        FlexAdVal2.TextMatrix(Linha, 3), FlexAdVal2.TextMatrix(Linha, 4), FlexAdVal2.TextMatrix(Linha, 5), FlexAdVal2.TextMatrix(Linha, 6), _
        FlexAdVal2.TextMatrix(Linha, 7), FlexAdVal2.TextMatrix(Linha, 8), xUsuario
        End If
    Next

'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO
'SAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPR
'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO

Limite = 40

    If Val(FlexAdVal2Total.TextMatrix(0, 0)) > Limite Then
        If ((Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) - Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)) > 0 Then
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) + 1
        Else
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)
        End If
    Else
    TotPag = 1
    End If
    Open SETIMPImpressoraPadrao For Output As #1
    
    Linha = 0
    Pag = 0
    
        For X = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        Linha = Linha + 1
        
        ZERO = Trim(Str(X))
        AWB = Trim(FlexAdVal2.TextMatrix(X, 2))
        CIA = Trim(FlexAdVal2.TextMatrix(X, 1))
        DESTINO = Trim(FlexAdVal2.TextMatrix(X, 3))
        SIGLADES = Trim(FlexAdVal2.TextMatrix(X, 4))
        VOLUME = Trim(FlexAdVal2.TextMatrix(X, 5))
        PESOREAL = Trim(FlexAdVal2.TextMatrix(X, 6))
        PESOTAX = Trim(FlexAdVal2.TextMatrix(X, 7))
        FRETE = Trim(FlexAdVal2.TextMatrix(X, 8))
        
        ZERO = Mid(ZERO, 1, 2)
        AWB = Mid(AWB, 1, 9)
        CIA = Mid(CIA, 1, 3)
        DESTINO = Mid(DESTINO, 1, 35)
        SIGLADES = Mid(SIGLADES, 1, 5)
        VOLUME = Mid(VOLUME, 1, 4)
        PESOREAL = Mid(PESOREAL, 1, 9)
        PESOTAX = Mid(PESOTAX, 1, 7)
        FRETE = Mid(FRETE, 1, 8)
        
        FRETE = Trim(Format(FRETE, "#0.00"))
        PESOREAL = Trim(Format(PESOREAL, "#0.00"))
        PESOTAX = Trim(Format(PESOTAX, "#0.00"))
        
        ZERO = ZERO & String(2 - Len(ZERO), " ")
        AWB = AWB & String(9 - Len(AWB), " ")
        CIA = CIA & String(3 - Len(CIA), " ")
        DESTINO = DESTINO & String(35 - Len(DESTINO), " ")
        SIGLADES = SIGLADES & String(5 - Len(SIGLADES), " ")
        VOLUME = String(4 - Len(VOLUME), " ") & VOLUME
        PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
        PESOTAX = String(7 - Len(PESOTAX), " ") & PESOTAX
        FRETE = String(8 - Len(FRETE), " ") & FRETE
        
        
            If Linha <= Limite And Linha <> 1 Then
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            ElseIf Linha = 1 Then
            Pag = Pag + 1
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(40) & "    I N T E C  T R A N S P O R T E S"
            Print #1, Chr(27) & "!" & Chr(86) & "Relacao de Cargas - TG30"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "Pagina " & Trim(Str(Pag)) & " de " & Trim(Str(TotPag))
            Print #1, Chr(27) & "!" & Chr(25) & "Data: " & Trim(Str(DataHora("DATA")))
            Print #1, Chr(27) & "!" & Chr(25) & "Hora: " & Trim(DataHora("HORA"))
            Print #1, Chr(27) & "!" & Chr(25) & "TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text)) & String(50 - Len("TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text))), " ") & Chr(27) & "!" & Chr(27) & "AD Valorem 1"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "   AWB        Cia  Destino                             Sigla  Vols  Peso Real  Peso Tx  Frete   "
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            'Else
            'Printer.NewPage
            'Linha = 0
            End If
        
        Next
        For Linha = X To Limite
        Print #1, Chr(27) & "!" & Chr(25) & "      "
        Next

Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"


ZERO = Trim(FlexAdVal2Total.TextMatrix(0, 0))
AWB = Trim(FlexAdVal2Total.TextMatrix(0, 2))
CIA = Trim(FlexAdVal2Total.TextMatrix(0, 1))
DESTINO = Trim(FlexAdVal2Total.TextMatrix(0, 3))
SIGLADES = Trim(FlexAdVal2Total.TextMatrix(0, 4))
VOLUME = Trim(FlexAdVal2Total.TextMatrix(0, 5))
PESOREAL = Trim(FlexAdVal2Total.TextMatrix(0, 6))
PESOTAX = Trim(FlexAdVal2Total.TextMatrix(0, 7))
FRETE = Trim(FlexAdVal2Total.TextMatrix(0, 8))

ZERO = Mid(AWB, 1, 2)
AWB = Mid(AWB, 1, 9)
CIA = Mid(CIA, 1, 3)
DESTINO = Mid(DESTINO, 1, 35)
SIGLADES = Mid(SIGLADES, 1, 5)
VOLUME = Mid(VOLUME, 1, 4)
PESOREAL = Mid(PESOREAL, 1, 9)
PESOTAX = Mid(PESOTAX, 1, 7)
FRETE = Mid(FRETE, 1, 8)

FRETE = Trim(Format(FRETE, "#0.00"))
PESOREAL = Trim(Format(PESOREAL, "#0.00"))
PESOTAX = Trim(Format(PESOTAX, "#0.00"))

ZERO = ZERO & String(9 - Len(ZERO), " ")
AWB = AWB & String(9 - Len(AWB), " ")
CIA = CIA & String(9 - Len(CIA), " ")
DESTINO = DESTINO & String(9 - Len(DESTINO), " ")
SIGLADES = SIGLADES & String(9 - Len(SIGLADES), " ")
VOLUME = String(9 - Len(VOLUME), " ") & VOLUME
PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
PESOTAX = String(9 - Len(PESOTAX), " ") & PESOTAX
FRETE = String(9 - Len(FRETE), " ") & FRETE

Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & "  " & PESOREAL & "  " & PESOTAX & "  " & FRETE
Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
Print #1, Chr(27) & "!" & Chr(27) & "A T E N C A O!   C A R G A   P E R C I V E L.   P R A Z O   D E   D U R A C A O   D E   4 8   H"
Print #1, Chr(27) & "!" & Chr(25) & "Recebido por:.......... Data: ......../......../........ Hora: .......:......."
Print #1, Chr(27) & "!" & Chr(25) & "Motorista: " & PriMaiuscula(TxtMotorista.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 1: " & PriMaiuscula(TxtAjudante1.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 2: " & PriMaiuscula(TxtAjudante2.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Carro: " & IIf(Len(Trim(TxtCodigo.Text)) > 0, "No. " & Trim(TxtCodigo.Text) & "  Placa: " & TxtPlaca.Text, TxtPlaca.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Emissor: "; UCase(xUsuario)
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Close #1
End If

End Sub

Sub AdValorem1NaoPerecivel()
'ADVALOREM 1 NAO PERECIVEL

FlexAdVal2.Clear
FlexAdVal2Total.Clear

FlexAdVal2.Rows = 2
FlexAdVal2Total.Rows = 1
FlexAdVal2.Cols = 11
FlexAdVal2Total.Cols = 11
FlexAdVal2.FixedRows = 1
FlexAdVal2.FixedCols = 0
FlexAdVal2Total.FixedRows = 0
FlexAdVal2Total.FixedCols = 0
FlexAdVal2.ColWidth(0) = 500
FlexAdVal2.ColWidth(1) = 500
FlexAdVal2.ColWidth(2) = 1000
FlexAdVal2.ColWidth(3) = 2000
FlexAdVal2.ColWidth(4) = 800
FlexAdVal2.ColWidth(5) = 800
FlexAdVal2.ColWidth(6) = 1000
FlexAdVal2.ColWidth(7) = 1000
FlexAdVal2.ColWidth(8) = 1000
FlexAdVal2.ColWidth(9) = 200
FlexAdVal2.ColWidth(10) = 200
FlexAdVal2Total.ColWidth(0) = 500
FlexAdVal2Total.ColWidth(1) = 500
FlexAdVal2Total.ColWidth(2) = 1000
FlexAdVal2Total.ColWidth(3) = 2000
FlexAdVal2Total.ColWidth(4) = 800
FlexAdVal2Total.ColWidth(5) = 800
FlexAdVal2Total.ColWidth(6) = 1000
FlexAdVal2Total.ColWidth(7) = 1000
FlexAdVal2Total.ColWidth(8) = 1000
FlexAdVal2Total.ColWidth(9) = 200
FlexAdVal2Total.ColWidth(10) = 200
FlexAdVal2.TextMatrix(0, 0) = "Filial"
FlexAdVal2.TextMatrix(0, 1) = "Cia."
FlexAdVal2.TextMatrix(0, 2) = "AWB"
FlexAdVal2.TextMatrix(0, 3) = "Destino"
FlexAdVal2.TextMatrix(0, 4) = "Sigla"
FlexAdVal2.TextMatrix(0, 5) = "Vols."
FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal2.TextMatrix(0, 8) = "Frete"
FlexAdVal2.TextMatrix(0, 9) = "P"
FlexAdVal2.TextMatrix(0, 10) = "A"

    For X = 1 To FlexAdVal1.Rows - 1
        If FlexAdVal1.TextMatrix(X, 9) = "" And Val(FlexAdVal1.TextMatrix(X, 10)) = 1 Then
        FlexAdVal2.Rows = FlexAdVal2.Rows + 1
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 0) = FlexAdVal1.TextMatrix(X, 0)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 1) = FlexAdVal1.TextMatrix(X, 1)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 2) = FlexAdVal1.TextMatrix(X, 2)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 3) = FlexAdVal1.TextMatrix(X, 3)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 4) = FlexAdVal1.TextMatrix(X, 4)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 5) = FlexAdVal1.TextMatrix(X, 5)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 6) = FlexAdVal1.TextMatrix(X, 6)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 7) = FlexAdVal1.TextMatrix(X, 7)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 8) = FlexAdVal1.TextMatrix(X, 8)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 9) = FlexAdVal1.TextMatrix(X, 9)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 10) = FlexAdVal1.TextMatrix(X, 10)
        End If
    Next
    
    xRec = 0
    xVols = 0
    xPesoReal = 0
    xPesoTaxado = 0
    xFrete = 0
        For X = 1 To FlexAdVal2.Rows - 1
            If Len(Trim(FlexAdVal2.TextMatrix(X, 0))) > 0 Then
            xRec = xRec + 1
            xVols = xVols + FlexAdVal2.TextMatrix(X, 5)
            xPesoReal = xPesoReal + FlexAdVal2.TextMatrix(X, 6)
            xPesoTaxado = xPesoTaxado + FlexAdVal2.TextMatrix(X, 7)
            xFrete = xFrete + FlexAdVal2.TextMatrix(X, 8)
            End If
        Next
    FlexAdVal2Total.TextMatrix(0, 0) = xRec
    FlexAdVal2Total.TextMatrix(0, 1) = ""
    FlexAdVal2Total.TextMatrix(0, 2) = ""
    FlexAdVal2Total.TextMatrix(0, 3) = ""
    FlexAdVal2Total.TextMatrix(0, 4) = ""
    FlexAdVal2Total.TextMatrix(0, 5) = xVols
    FlexAdVal2Total.TextMatrix(0, 6) = xPesoReal
    FlexAdVal2Total.TextMatrix(0, 7) = xPesoTaxado
    FlexAdVal2Total.TextMatrix(0, 8) = xFrete

If Val(FlexAdVal2Total.TextMatrix(0, 0)) > 0 Then
If de_informa.rsCapturaManifesto.State = 1 Then de_informa.rsCapturaManifesto.Close
de_informa.CapturaManifesto Mid(TxtNomeFilial.Text, 1, 2)

    If de_informa.rsCapturaManifesto.RecordCount > 0 Then
        If Not IsNull(de_informa.rsCapturaManifesto.Fields("manifesto")) Then
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))), "0") & Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))
        Else
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
        End If
    Else
    xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
    End If
    
    For Linha = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        If Val(FlexAdVal2.TextMatrix(Linha, 0)) > 0 Then
        xCodAwb = FlexAdVal2.TextMatrix(Linha, 0) & FlexAdVal2.TextMatrix(Linha, 1) & String(10 - Len(Trim(Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2))), "0") & Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2) & Mid(FlexAdVal2.TextMatrix(Linha, 2), Len(FlexAdVal2.TextMatrix(Linha, 2)))
        de_informa.InsAirManifesto xFilialManifesto, FlexAdVal2.TextMatrix(Linha, 0), Trim(Str(Val(Mid(xFilialManifesto, 3)))), DataHora("DATA"), _
        DataHora("HORA"), FlexAdVal2.TextMatrix(Linha, 10), FlexAdVal2.TextMatrix(Linha, 9), FlexAdVal2.TextMatrix(Linha, 1), UCase(TxtNomeCia.Text), xCodAwb, UCase(TxtAeroporto.Text), TxtSiglaOrigem.Text, _
        FlexAdVal2.TextMatrix(Linha, 3), FlexAdVal2.TextMatrix(Linha, 4), FlexAdVal2.TextMatrix(Linha, 5), FlexAdVal2.TextMatrix(Linha, 6), _
        FlexAdVal2.TextMatrix(Linha, 7), FlexAdVal2.TextMatrix(Linha, 8), xUsuario
        End If
    Next

'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO
'SAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPR
'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO

Limite = 40

    If Val(FlexAdVal2Total.TextMatrix(0, 0)) > Limite Then
        If ((Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) - Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)) > 0 Then
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) + 1
        Else
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)
        End If
    Else
    TotPag = 1
    End If
    Open SETIMPImpressoraPadrao For Output As #1
    
    Linha = 0
    Pag = 0
    
        For X = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        Linha = Linha + 1
        
        ZERO = Trim(Str(X))
        AWB = Trim(FlexAdVal2.TextMatrix(X, 2))
        CIA = Trim(FlexAdVal2.TextMatrix(X, 1))
        DESTINO = Trim(FlexAdVal2.TextMatrix(X, 3))
        SIGLADES = Trim(FlexAdVal2.TextMatrix(X, 4))
        VOLUME = Trim(FlexAdVal2.TextMatrix(X, 5))
        PESOREAL = Trim(FlexAdVal2.TextMatrix(X, 6))
        PESOTAX = Trim(FlexAdVal2.TextMatrix(X, 7))
        FRETE = Trim(FlexAdVal2.TextMatrix(X, 8))
        
        ZERO = Mid(ZERO, 1, 2)
        AWB = Mid(AWB, 1, 9)
        CIA = Mid(CIA, 1, 3)
        DESTINO = Mid(DESTINO, 1, 35)
        SIGLADES = Mid(SIGLADES, 1, 5)
        VOLUME = Mid(VOLUME, 1, 4)
        PESOREAL = Mid(PESOREAL, 1, 9)
        PESOTAX = Mid(PESOTAX, 1, 7)
        FRETE = Mid(FRETE, 1, 8)
        
        FRETE = Trim(Format(FRETE, "#0.00"))
        PESOREAL = Trim(Format(PESOREAL, "#0.00"))
        PESOTAX = Trim(Format(PESOTAX, "#0.00"))
        
        ZERO = ZERO & String(2 - Len(ZERO), " ")
        AWB = AWB & String(9 - Len(AWB), " ")
        CIA = CIA & String(3 - Len(CIA), " ")
        DESTINO = DESTINO & String(35 - Len(DESTINO), " ")
        SIGLADES = SIGLADES & String(5 - Len(SIGLADES), " ")
        VOLUME = String(4 - Len(VOLUME), " ") & VOLUME
        PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
        PESOTAX = String(7 - Len(PESOTAX), " ") & PESOTAX
        FRETE = String(8 - Len(FRETE), " ") & FRETE
        
        
            If Linha <= Limite And Linha <> 1 Then
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            ElseIf Linha = 1 Then
            Pag = Pag + 1
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(40) & "    I N T E C  T R A N S P O R T E S"
            Print #1, Chr(27) & "!" & Chr(86) & "Relacao de Cargas - TG30"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "Pagina " & Trim(Str(Pag)) & " de " & Trim(Str(TotPag))
            Print #1, Chr(27) & "!" & Chr(25) & "Data: " & Trim(Str(DataHora("DATA")))
            Print #1, Chr(27) & "!" & Chr(25) & "Hora: " & Trim(DataHora("HORA"))
            Print #1, Chr(27) & "!" & Chr(25) & "TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text)) & String(50 - Len("TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text))), " ") & Chr(27) & "!" & Chr(27) & "AD Valorem 1"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "   AWB        Cia  Destino                             Sigla  Vols  Peso Real  Peso Tx  Frete   "
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            'Else
            'Printer.NewPage
            'Linha = 0
            End If
        
        Next
        For Linha = X To Limite
        Print #1, Chr(27) & "!" & Chr(25) & "      "
        Next

Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"


ZERO = Trim(FlexAdVal2Total.TextMatrix(0, 0))
AWB = Trim(FlexAdVal2Total.TextMatrix(0, 2))
CIA = Trim(FlexAdVal2Total.TextMatrix(0, 1))
DESTINO = Trim(FlexAdVal2Total.TextMatrix(0, 3))
SIGLADES = Trim(FlexAdVal2Total.TextMatrix(0, 4))
VOLUME = Trim(FlexAdVal2Total.TextMatrix(0, 5))
PESOREAL = Trim(FlexAdVal2Total.TextMatrix(0, 6))
PESOTAX = Trim(FlexAdVal2Total.TextMatrix(0, 7))
FRETE = Trim(FlexAdVal2Total.TextMatrix(0, 8))

ZERO = Mid(AWB, 1, 2)
AWB = Mid(AWB, 1, 9)
CIA = Mid(CIA, 1, 3)
DESTINO = Mid(DESTINO, 1, 35)
SIGLADES = Mid(SIGLADES, 1, 5)
VOLUME = Mid(VOLUME, 1, 4)
PESOREAL = Mid(PESOREAL, 1, 9)
PESOTAX = Mid(PESOTAX, 1, 7)
FRETE = Mid(FRETE, 1, 8)

FRETE = Trim(Format(FRETE, "#0.00"))
PESOREAL = Trim(Format(PESOREAL, "#0.00"))
PESOTAX = Trim(Format(PESOTAX, "#0.00"))

ZERO = ZERO & String(9 - Len(ZERO), " ")
AWB = AWB & String(9 - Len(AWB), " ")
CIA = CIA & String(9 - Len(CIA), " ")
DESTINO = DESTINO & String(9 - Len(DESTINO), " ")
SIGLADES = SIGLADES & String(9 - Len(SIGLADES), " ")
VOLUME = String(9 - Len(VOLUME), " ") & VOLUME
PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
PESOTAX = String(9 - Len(PESOTAX), " ") & PESOTAX
FRETE = String(9 - Len(FRETE), " ") & FRETE

Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & "  " & PESOREAL & "  " & PESOTAX & "  " & FRETE
Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
Print #1, Chr(27) & "!" & Chr(25) & "      "
Print #1, Chr(27) & "!" & Chr(25) & "Recebido por:.......... Data: ......../......../........ Hora: .......:......."
Print #1, Chr(27) & "!" & Chr(25) & "Motorista: " & PriMaiuscula(TxtMotorista.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 1: " & PriMaiuscula(TxtAjudante1.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 2: " & PriMaiuscula(TxtAjudante2.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Carro: " & IIf(Len(Trim(TxtCodigo.Text)) > 0, "No. " & Trim(TxtCodigo.Text) & "  Placa: " & TxtPlaca.Text, TxtPlaca.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Emissor: "; UCase(xUsuario)
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Close #1
    End If
End Sub

Sub AdValorem2NaoPerecivel()
'ADVALOREM 2 NAO PERECIVEL

FlexAdVal2.Clear
FlexAdVal2Total.Clear

FlexAdVal2.Rows = 2
FlexAdVal2Total.Rows = 1
FlexAdVal2.Cols = 11
FlexAdVal2Total.Cols = 11
FlexAdVal2.FixedRows = 1
FlexAdVal2.FixedCols = 0
FlexAdVal2Total.FixedRows = 0
FlexAdVal2Total.FixedCols = 0
FlexAdVal2.ColWidth(0) = 500
FlexAdVal2.ColWidth(1) = 500
FlexAdVal2.ColWidth(2) = 1000
FlexAdVal2.ColWidth(3) = 2000
FlexAdVal2.ColWidth(4) = 800
FlexAdVal2.ColWidth(5) = 800
FlexAdVal2.ColWidth(6) = 1000
FlexAdVal2.ColWidth(7) = 1000
FlexAdVal2.ColWidth(8) = 1000
FlexAdVal2.ColWidth(9) = 200
FlexAdVal2.ColWidth(10) = 200
FlexAdVal2Total.ColWidth(0) = 500
FlexAdVal2Total.ColWidth(1) = 500
FlexAdVal2Total.ColWidth(2) = 1000
FlexAdVal2Total.ColWidth(3) = 2000
FlexAdVal2Total.ColWidth(4) = 800
FlexAdVal2Total.ColWidth(5) = 800
FlexAdVal2Total.ColWidth(6) = 1000
FlexAdVal2Total.ColWidth(7) = 1000
FlexAdVal2Total.ColWidth(8) = 1000
FlexAdVal2Total.ColWidth(9) = 200
FlexAdVal2Total.ColWidth(10) = 200
FlexAdVal2.TextMatrix(0, 0) = "Filial"
FlexAdVal2.TextMatrix(0, 1) = "Cia."
FlexAdVal2.TextMatrix(0, 2) = "AWB"
FlexAdVal2.TextMatrix(0, 3) = "Destino"
FlexAdVal2.TextMatrix(0, 4) = "Sigla"
FlexAdVal2.TextMatrix(0, 5) = "Vols."
FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal2.TextMatrix(0, 8) = "Frete"
FlexAdVal2.TextMatrix(0, 9) = "P"
FlexAdVal2.TextMatrix(0, 10) = "A"

    For X = 1 To FlexAdVal1.Rows - 1
        If FlexAdVal1.TextMatrix(X, 9) = "" And Val(FlexAdVal1.TextMatrix(X, 10)) = 2 Then
        FlexAdVal2.Rows = FlexAdVal2.Rows + 1
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 0) = FlexAdVal1.TextMatrix(X, 0)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 1) = FlexAdVal1.TextMatrix(X, 1)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 2) = FlexAdVal1.TextMatrix(X, 2)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 3) = FlexAdVal1.TextMatrix(X, 3)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 4) = FlexAdVal1.TextMatrix(X, 4)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 5) = FlexAdVal1.TextMatrix(X, 5)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 6) = FlexAdVal1.TextMatrix(X, 6)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 7) = FlexAdVal1.TextMatrix(X, 7)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 8) = FlexAdVal1.TextMatrix(X, 8)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 9) = FlexAdVal1.TextMatrix(X, 9)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 10) = FlexAdVal1.TextMatrix(X, 10)
        End If
    Next
    
    xRec = 0
    xVols = 0
    xPesoReal = 0
    xPesoTaxado = 0
    xFrete = 0
        For X = 1 To FlexAdVal2.Rows - 1
            If Len(Trim(FlexAdVal2.TextMatrix(X, 0))) > 0 Then
            xRec = xRec + 1
            xVols = xVols + FlexAdVal2.TextMatrix(X, 5)
            xPesoReal = xPesoReal + FlexAdVal2.TextMatrix(X, 6)
            xPesoTaxado = xPesoTaxado + FlexAdVal2.TextMatrix(X, 7)
            xFrete = xFrete + FlexAdVal2.TextMatrix(X, 8)
            End If
        Next
    FlexAdVal2Total.TextMatrix(0, 0) = xRec
    FlexAdVal2Total.TextMatrix(0, 1) = ""
    FlexAdVal2Total.TextMatrix(0, 2) = ""
    FlexAdVal2Total.TextMatrix(0, 3) = ""
    FlexAdVal2Total.TextMatrix(0, 4) = ""
    FlexAdVal2Total.TextMatrix(0, 5) = xVols
    FlexAdVal2Total.TextMatrix(0, 6) = xPesoReal
    FlexAdVal2Total.TextMatrix(0, 7) = xPesoTaxado
    FlexAdVal2Total.TextMatrix(0, 8) = xFrete

If Val(FlexAdVal2Total.TextMatrix(0, 0)) > 0 Then
If de_informa.rsCapturaManifesto.State = 1 Then de_informa.rsCapturaManifesto.Close
de_informa.CapturaManifesto Mid(TxtNomeFilial.Text, 1, 2)

    If de_informa.rsCapturaManifesto.RecordCount > 0 Then
        If Not IsNull(de_informa.rsCapturaManifesto.Fields("manifesto")) Then
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))), "0") & Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))
        Else
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
        End If
    Else
    xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
    End If
    
    For Linha = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        If Val(FlexAdVal2.TextMatrix(Linha, 0)) > 0 Then
        xCodAwb = FlexAdVal2.TextMatrix(Linha, 0) & FlexAdVal2.TextMatrix(Linha, 1) & String(10 - Len(Trim(Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2))), "0") & Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2) & Mid(FlexAdVal2.TextMatrix(Linha, 2), Len(FlexAdVal2.TextMatrix(Linha, 2)))
        de_informa.InsAirManifesto xFilialManifesto, FlexAdVal2.TextMatrix(Linha, 0), Trim(Str(Val(Mid(xFilialManifesto, 3)))), DataHora("DATA"), _
        DataHora("HORA"), FlexAdVal2.TextMatrix(Linha, 10), FlexAdVal2.TextMatrix(Linha, 9), FlexAdVal2.TextMatrix(Linha, 1), UCase(TxtNomeCia.Text), xCodAwb, UCase(TxtAeroporto.Text), TxtSiglaOrigem.Text, _
        FlexAdVal2.TextMatrix(Linha, 3), FlexAdVal2.TextMatrix(Linha, 4), FlexAdVal2.TextMatrix(Linha, 5), FlexAdVal2.TextMatrix(Linha, 6), _
        FlexAdVal2.TextMatrix(Linha, 7), FlexAdVal2.TextMatrix(Linha, 8), xUsuario
        End If
    Next

'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO
'SAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPR
'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO

Limite = 40

    If Val(FlexAdVal2Total.TextMatrix(0, 0)) > Limite Then
        If ((Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) - Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)) > 0 Then
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) + 1
        Else
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)
        End If
    Else
    TotPag = 1
    End If
    Open SETIMPImpressoraPadrao For Output As #1
    
    Linha = 0
    Pag = 0
    
        For X = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        Linha = Linha + 1
        
        ZERO = Trim(Str(X))
        AWB = Trim(FlexAdVal2.TextMatrix(X, 2))
        CIA = Trim(FlexAdVal2.TextMatrix(X, 1))
        DESTINO = Trim(FlexAdVal2.TextMatrix(X, 3))
        SIGLADES = Trim(FlexAdVal2.TextMatrix(X, 4))
        VOLUME = Trim(FlexAdVal2.TextMatrix(X, 5))
        PESOREAL = Trim(FlexAdVal2.TextMatrix(X, 6))
        PESOTAX = Trim(FlexAdVal2.TextMatrix(X, 7))
        FRETE = Trim(FlexAdVal2.TextMatrix(X, 8))
        
        ZERO = Mid(ZERO, 1, 2)
        AWB = Mid(AWB, 1, 9)
        CIA = Mid(CIA, 1, 3)
        DESTINO = Mid(DESTINO, 1, 35)
        SIGLADES = Mid(SIGLADES, 1, 5)
        VOLUME = Mid(VOLUME, 1, 4)
        PESOREAL = Mid(PESOREAL, 1, 9)
        PESOTAX = Mid(PESOTAX, 1, 7)
        FRETE = Mid(FRETE, 1, 8)
        
        FRETE = Trim(Format(FRETE, "#0.00"))
        PESOREAL = Trim(Format(PESOREAL, "#0.00"))
        PESOTAX = Trim(Format(PESOTAX, "#0.00"))
        
        ZERO = ZERO & String(2 - Len(ZERO), " ")
        AWB = AWB & String(9 - Len(AWB), " ")
        CIA = CIA & String(3 - Len(CIA), " ")
        DESTINO = DESTINO & String(35 - Len(DESTINO), " ")
        SIGLADES = SIGLADES & String(5 - Len(SIGLADES), " ")
        VOLUME = String(4 - Len(VOLUME), " ") & VOLUME
        PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
        PESOTAX = String(7 - Len(PESOTAX), " ") & PESOTAX
        FRETE = String(8 - Len(FRETE), " ") & FRETE
        
        
            If Linha <= Limite And Linha <> 1 Then
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            ElseIf Linha = 1 Then
            Pag = Pag + 1
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(40) & "    I N T E C  T R A N S P O R T E S"
            Print #1, Chr(27) & "!" & Chr(86) & "Relacao de Cargas - TG30"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "Pagina " & Trim(Str(Pag)) & " de " & Trim(Str(TotPag))
            Print #1, Chr(27) & "!" & Chr(25) & "Data: " & Trim(Str(DataHora("DATA")))
            Print #1, Chr(27) & "!" & Chr(25) & "Hora: " & Trim(DataHora("HORA"))
            Print #1, Chr(27) & "!" & Chr(25) & "TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text)) & String(50 - Len("TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text))), " ") & Chr(27) & "!" & Chr(27) & "AD Valorem 2"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "   AWB        Cia  Destino                             Sigla  Vols  Peso Real  Peso Tx  Frete   "
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            'Else
            'Printer.NewPage
            'Linha = 0
            End If
        
        Next
        For Linha = X To Limite
        Print #1, Chr(27) & "!" & Chr(25) & "      "
        Next

Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"


ZERO = Trim(FlexAdVal2Total.TextMatrix(0, 0))
AWB = Trim(FlexAdVal2Total.TextMatrix(0, 2))
CIA = Trim(FlexAdVal2Total.TextMatrix(0, 1))
DESTINO = Trim(FlexAdVal2Total.TextMatrix(0, 3))
SIGLADES = Trim(FlexAdVal2Total.TextMatrix(0, 4))
VOLUME = Trim(FlexAdVal2Total.TextMatrix(0, 5))
PESOREAL = Trim(FlexAdVal2Total.TextMatrix(0, 6))
PESOTAX = Trim(FlexAdVal2Total.TextMatrix(0, 7))
FRETE = Trim(FlexAdVal2Total.TextMatrix(0, 8))

ZERO = Mid(AWB, 1, 2)
AWB = Mid(AWB, 1, 9)
CIA = Mid(CIA, 1, 3)
DESTINO = Mid(DESTINO, 1, 35)
SIGLADES = Mid(SIGLADES, 1, 5)
VOLUME = Mid(VOLUME, 1, 4)
PESOREAL = Mid(PESOREAL, 1, 9)
PESOTAX = Mid(PESOTAX, 1, 7)
FRETE = Mid(FRETE, 1, 8)

FRETE = Trim(Format(FRETE, "#0.00"))
PESOREAL = Trim(Format(PESOREAL, "#0.00"))
PESOTAX = Trim(Format(PESOTAX, "#0.00"))

ZERO = ZERO & String(9 - Len(ZERO), " ")
AWB = AWB & String(9 - Len(AWB), " ")
CIA = CIA & String(9 - Len(CIA), " ")
DESTINO = DESTINO & String(9 - Len(DESTINO), " ")
SIGLADES = SIGLADES & String(9 - Len(SIGLADES), " ")
VOLUME = String(9 - Len(VOLUME), " ") & VOLUME
PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
PESOTAX = String(9 - Len(PESOTAX), " ") & PESOTAX
FRETE = String(9 - Len(FRETE), " ") & FRETE

Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & "  " & PESOREAL & "  " & PESOTAX & "  " & FRETE
Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
Print #1, Chr(27) & "!" & Chr(25) & "      "
Print #1, Chr(27) & "!" & Chr(25) & "Recebido por:.......... Data: ......../......../........ Hora: .......:......."
Print #1, Chr(27) & "!" & Chr(25) & "Motorista: " & PriMaiuscula(TxtMotorista.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 1: " & PriMaiuscula(TxtAjudante1.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 2: " & PriMaiuscula(TxtAjudante2.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Carro: " & IIf(Len(Trim(TxtCodigo.Text)) > 0, "No. " & Trim(TxtCodigo.Text) & "  Placa: " & TxtPlaca.Text, TxtPlaca.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Emissor: "; UCase(xUsuario)
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Close #1
    End If
End Sub

Sub AdValorem2Perecivel()
'ADVALOREM 2 PERECIVEL

FlexAdVal2.Clear
FlexAdVal2Total.Clear

FlexAdVal2.Rows = 2
FlexAdVal2Total.Rows = 1
FlexAdVal2.Cols = 11
FlexAdVal2Total.Cols = 11
FlexAdVal2.FixedRows = 1
FlexAdVal2.FixedCols = 0
FlexAdVal2Total.FixedRows = 0
FlexAdVal2Total.FixedCols = 0
FlexAdVal2.ColWidth(0) = 500
FlexAdVal2.ColWidth(1) = 500
FlexAdVal2.ColWidth(2) = 1000
FlexAdVal2.ColWidth(3) = 2000
FlexAdVal2.ColWidth(4) = 800
FlexAdVal2.ColWidth(5) = 800
FlexAdVal2.ColWidth(6) = 1000
FlexAdVal2.ColWidth(7) = 1000
FlexAdVal2.ColWidth(8) = 1000
FlexAdVal2.ColWidth(9) = 200
FlexAdVal2.ColWidth(10) = 200
FlexAdVal2Total.ColWidth(0) = 500
FlexAdVal2Total.ColWidth(1) = 500
FlexAdVal2Total.ColWidth(2) = 1000
FlexAdVal2Total.ColWidth(3) = 2000
FlexAdVal2Total.ColWidth(4) = 800
FlexAdVal2Total.ColWidth(5) = 800
FlexAdVal2Total.ColWidth(6) = 1000
FlexAdVal2Total.ColWidth(7) = 1000
FlexAdVal2Total.ColWidth(8) = 1000
FlexAdVal2Total.ColWidth(9) = 200
FlexAdVal2Total.ColWidth(10) = 200
FlexAdVal2.TextMatrix(0, 0) = "Filial"
FlexAdVal2.TextMatrix(0, 1) = "Cia."
FlexAdVal2.TextMatrix(0, 2) = "AWB"
FlexAdVal2.TextMatrix(0, 3) = "Destino"
FlexAdVal2.TextMatrix(0, 4) = "Sigla"
FlexAdVal2.TextMatrix(0, 5) = "Vols."
FlexAdVal2.TextMatrix(0, 6) = "Peso Real"
FlexAdVal2.TextMatrix(0, 7) = "Peso Tax."
FlexAdVal2.TextMatrix(0, 8) = "Frete"
FlexAdVal2.TextMatrix(0, 9) = "P"
FlexAdVal2.TextMatrix(0, 10) = "A"

    For X = 1 To FlexAdVal1.Rows - 1
        If FlexAdVal1.TextMatrix(X, 9) = "S" And Val(FlexAdVal1.TextMatrix(X, 10)) = 2 Then
        FlexAdVal2.Rows = FlexAdVal2.Rows + 1
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 0) = FlexAdVal1.TextMatrix(X, 0)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 1) = FlexAdVal1.TextMatrix(X, 1)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 2) = FlexAdVal1.TextMatrix(X, 2)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 3) = FlexAdVal1.TextMatrix(X, 3)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 4) = FlexAdVal1.TextMatrix(X, 4)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 5) = FlexAdVal1.TextMatrix(X, 5)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 6) = FlexAdVal1.TextMatrix(X, 6)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 7) = FlexAdVal1.TextMatrix(X, 7)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 8) = FlexAdVal1.TextMatrix(X, 8)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 9) = FlexAdVal1.TextMatrix(X, 9)
        FlexAdVal2.TextMatrix(FlexAdVal2.Rows - 2, 10) = FlexAdVal1.TextMatrix(X, 10)
        End If
    Next
    
    xRec = 0
    xVols = 0
    xPesoReal = 0
    xPesoTaxado = 0
    xFrete = 0
        For X = 1 To FlexAdVal2.Rows - 1
            If Len(Trim(FlexAdVal2.TextMatrix(X, 0))) > 0 Then
            xRec = xRec + 1
            xVols = xVols + FlexAdVal2.TextMatrix(X, 5)
            xPesoReal = xPesoReal + FlexAdVal2.TextMatrix(X, 6)
            xPesoTaxado = xPesoTaxado + FlexAdVal2.TextMatrix(X, 7)
            xFrete = xFrete + FlexAdVal2.TextMatrix(X, 8)
            End If
        Next
    FlexAdVal2Total.TextMatrix(0, 0) = xRec
    FlexAdVal2Total.TextMatrix(0, 1) = ""
    FlexAdVal2Total.TextMatrix(0, 2) = ""
    FlexAdVal2Total.TextMatrix(0, 3) = ""
    FlexAdVal2Total.TextMatrix(0, 4) = ""
    FlexAdVal2Total.TextMatrix(0, 5) = xVols
    FlexAdVal2Total.TextMatrix(0, 6) = xPesoReal
    FlexAdVal2Total.TextMatrix(0, 7) = xPesoTaxado
    FlexAdVal2Total.TextMatrix(0, 8) = xFrete

If Val(FlexAdVal2Total.TextMatrix(0, 0)) > 0 Then
If de_informa.rsCapturaManifesto.State = 1 Then de_informa.rsCapturaManifesto.Close
de_informa.CapturaManifesto Mid(TxtNomeFilial.Text, 1, 2)

    If de_informa.rsCapturaManifesto.RecordCount > 0 Then
        If Not IsNull(de_informa.rsCapturaManifesto.Fields("manifesto")) Then
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))), "0") & Trim(Str(Val(de_informa.rsCapturaManifesto.Fields("manifesto")) + 1))
        Else
        xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
        End If
    Else
    xFilialManifesto = Mid(TxtNomeFilial.Text, 1, 2) & String(10 - Len(Trim(Str(Val(0) + 1))), "0") & Trim(Str(Val(0) + 1))
    End If
    
    For Linha = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        If Val(FlexAdVal2.TextMatrix(Linha, 0)) > 0 Then
        xCodAwb = FlexAdVal2.TextMatrix(Linha, 0) & FlexAdVal2.TextMatrix(Linha, 1) & String(10 - Len(Trim(Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2))), "0") & Mid(FlexAdVal2.TextMatrix(Linha, 2), 1, Len(FlexAdVal2.TextMatrix(Linha, 2)) - 2) & Mid(FlexAdVal2.TextMatrix(Linha, 2), Len(FlexAdVal2.TextMatrix(Linha, 2)))
        de_informa.InsAirManifesto xFilialManifesto, FlexAdVal2.TextMatrix(Linha, 0), Trim(Str(Val(Mid(xFilialManifesto, 3)))), DataHora("DATA"), _
        DataHora("HORA"), FlexAdVal2.TextMatrix(Linha, 10), FlexAdVal2.TextMatrix(Linha, 9), FlexAdVal2.TextMatrix(Linha, 1), UCase(TxtNomeCia.Text), xCodAwb, UCase(TxtAeroporto.Text), TxtSiglaOrigem.Text, _
        FlexAdVal2.TextMatrix(Linha, 3), FlexAdVal2.TextMatrix(Linha, 4), FlexAdVal2.TextMatrix(Linha, 5), FlexAdVal2.TextMatrix(Linha, 6), _
        FlexAdVal2.TextMatrix(Linha, 7), FlexAdVal2.TextMatrix(Linha, 8), xUsuario
        End If
    Next

'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO
'SAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPR
'IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO  IMPRESSAO

Limite = 40

    If Val(FlexAdVal2Total.TextMatrix(0, 0)) > Limite Then
        If ((Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) - Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)) > 0 Then
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite) + 1
        Else
        TotPag = Int(Val(FlexAdVal2Total.TextMatrix(0, 0)) / Limite)
        End If
    Else
    TotPag = 1
    End If
    Open SETIMPImpressoraPadrao For Output As #1
    
    Linha = 0
    Pag = 0
    
        For X = 1 To Val(FlexAdVal2Total.TextMatrix(0, 0))
        Linha = Linha + 1
        
        ZERO = Trim(Str(X))
        AWB = Trim(FlexAdVal2.TextMatrix(X, 2))
        CIA = Trim(FlexAdVal2.TextMatrix(X, 1))
        DESTINO = Trim(FlexAdVal2.TextMatrix(X, 3))
        SIGLADES = Trim(FlexAdVal2.TextMatrix(X, 4))
        VOLUME = Trim(FlexAdVal2.TextMatrix(X, 5))
        PESOREAL = Trim(FlexAdVal2.TextMatrix(X, 6))
        PESOTAX = Trim(FlexAdVal2.TextMatrix(X, 7))
        FRETE = Trim(FlexAdVal2.TextMatrix(X, 8))
        
        ZERO = Mid(ZERO, 1, 2)
        AWB = Mid(AWB, 1, 9)
        CIA = Mid(CIA, 1, 3)
        DESTINO = Mid(DESTINO, 1, 35)
        SIGLADES = Mid(SIGLADES, 1, 5)
        VOLUME = Mid(VOLUME, 1, 4)
        PESOREAL = Mid(PESOREAL, 1, 9)
        PESOTAX = Mid(PESOTAX, 1, 7)
        FRETE = Mid(FRETE, 1, 8)
        
        FRETE = Trim(Format(FRETE, "#0.00"))
        PESOREAL = Trim(Format(PESOREAL, "#0.00"))
        PESOTAX = Trim(Format(PESOTAX, "#0.00"))
        
        ZERO = ZERO & String(2 - Len(ZERO), " ")
        AWB = AWB & String(9 - Len(AWB), " ")
        CIA = CIA & String(3 - Len(CIA), " ")
        DESTINO = DESTINO & String(35 - Len(DESTINO), " ")
        SIGLADES = SIGLADES & String(5 - Len(SIGLADES), " ")
        VOLUME = String(4 - Len(VOLUME), " ") & VOLUME
        PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
        PESOTAX = String(7 - Len(PESOTAX), " ") & PESOTAX
        FRETE = String(8 - Len(FRETE), " ") & FRETE
        
        
            If Linha <= Limite And Linha <> 1 Then
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            ElseIf Linha = 1 Then
            Pag = Pag + 1
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(40) & "    I N T E C  T R A N S P O R T E S"
            Print #1, Chr(27) & "!" & Chr(86) & "Relacao de Cargas - TG30"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "Pagina " & Trim(Str(Pag)) & " de " & Trim(Str(TotPag))
            Print #1, Chr(27) & "!" & Chr(25) & "Data: " & Trim(Str(DataHora("DATA")))
            Print #1, Chr(27) & "!" & Chr(25) & "Hora: " & Trim(DataHora("HORA"))
            Print #1, Chr(27) & "!" & Chr(25) & "TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text)) & String(50 - Len("TG30  Numero: " & Mid(xFilialManifesto, 1, 2) & "-" & Mid(xFilialManifesto, 3) & "  -  " & UCase(Trim(TxtNomeCia.Text))), " ") & Chr(27) & "!" & Chr(27) & "AD Valorem 2"
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & "   AWB        Cia  Destino                             Sigla  Vols  Peso Real  Peso Tx  Frete   "
            Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
            Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & " " & PESOREAL & " " & PESOTAX & "  " & FRETE
            'Else
            'Printer.NewPage
            'Linha = 0
            End If
        
        Next
        For Linha = X To Limite
        Print #1, Chr(27) & "!" & Chr(25) & "      "
        Next

Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"


ZERO = Trim(FlexAdVal2Total.TextMatrix(0, 0))
AWB = Trim(FlexAdVal2Total.TextMatrix(0, 2))
CIA = Trim(FlexAdVal2Total.TextMatrix(0, 1))
DESTINO = Trim(FlexAdVal2Total.TextMatrix(0, 3))
SIGLADES = Trim(FlexAdVal2Total.TextMatrix(0, 4))
VOLUME = Trim(FlexAdVal2Total.TextMatrix(0, 5))
PESOREAL = Trim(FlexAdVal2Total.TextMatrix(0, 6))
PESOTAX = Trim(FlexAdVal2Total.TextMatrix(0, 7))
FRETE = Trim(FlexAdVal2Total.TextMatrix(0, 8))

ZERO = Mid(AWB, 1, 2)
AWB = Mid(AWB, 1, 9)
CIA = Mid(CIA, 1, 3)
DESTINO = Mid(DESTINO, 1, 35)
SIGLADES = Mid(SIGLADES, 1, 5)
VOLUME = Mid(VOLUME, 1, 4)
PESOREAL = Mid(PESOREAL, 1, 9)
PESOTAX = Mid(PESOTAX, 1, 7)
FRETE = Mid(FRETE, 1, 8)

FRETE = Trim(Format(FRETE, "#0.00"))
PESOREAL = Trim(Format(PESOREAL, "#0.00"))
PESOTAX = Trim(Format(PESOTAX, "#0.00"))

ZERO = ZERO & String(9 - Len(ZERO), " ")
AWB = AWB & String(9 - Len(AWB), " ")
CIA = CIA & String(9 - Len(CIA), " ")
DESTINO = DESTINO & String(9 - Len(DESTINO), " ")
SIGLADES = SIGLADES & String(9 - Len(SIGLADES), " ")
VOLUME = String(9 - Len(VOLUME), " ") & VOLUME
PESOREAL = String(9 - Len(PESOREAL), " ") & PESOREAL
PESOTAX = String(9 - Len(PESOTAX), " ") & PESOTAX
FRETE = String(9 - Len(FRETE), " ") & FRETE

Print #1, Chr(27) & "!" & Chr(25) & ZERO & " " & AWB & "  " & CIA & "  " & DESTINO & "  " & SIGLADES & "  " & VOLUME & "  " & PESOREAL & "  " & PESOTAX & "  " & FRETE
Print #1, Chr(27) & "!" & Chr(25) & "================================================================================================"
Print #1, Chr(27) & "!" & Chr(27) & "A T E N C A O!   C A R G A   P E R C I V E L.   P R A Z O   D E   D U R A C A O   D E   4 8   H"
Print #1, Chr(27) & "!" & Chr(25) & "Recebido por:.......... Data: ......../......../........ Hora: .......:......."
Print #1, Chr(27) & "!" & Chr(25) & "Motorista: " & PriMaiuscula(TxtMotorista.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 1: " & PriMaiuscula(TxtAjudante1.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Ajudante 2: " & PriMaiuscula(TxtAjudante2.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Carro: " & IIf(Len(Trim(TxtCodigo.Text)) > 0, "No. " & Trim(TxtCodigo.Text) & "  Placa: " & TxtPlaca.Text, TxtPlaca.Text)
Print #1, Chr(27) & "!" & Chr(25) & "Emissor: "; UCase(xUsuario)
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Close #1
End If

End Sub
