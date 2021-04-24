VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAIRAcompanha 
   Caption         =   "Informa - Acompanhamento de Cliente"
   ClientHeight    =   8295
   ClientLeft      =   720
   ClientTop       =   1665
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmAIRAcompanha.frx":0000
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
      Height          =   6555
      Left            =   120
      TabIndex        =   27
      Top             =   1620
      Width           =   11775
      Begin VB.CheckBox ChkIncluir 
         Caption         =   "Incluir AWBs Já Informados por E-Mail"
         Height          =   195
         Left            =   8580
         TabIndex        =   29
         Top             =   660
         Width           =   3015
      End
      Begin VB.CommandButton CmdEMAIL 
         Caption         =   "Enviar E-Mails para Representantes"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3060
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdGerarArq 
         Caption         =   "Geração de Arquivo de Movimento"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8580
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmdInserirVOO 
         Caption         =   "Inserir / Alterar Vôo deste AWB"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid FlexAWB 
         Height          =   5535
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9763
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
         TabIndex        =   28
         Top             =   480
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
      TabIndex        =   18
      Top             =   120
      Width           =   6135
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   4800
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   2835
         Begin VB.OptionButton Option5 
            Caption         =   "AWBs"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            ToolTipText     =   "Traz Somente AWBs"
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton OptNF 
            Caption         =   "NF"
            Height          =   195
            Left            =   1860
            TabIndex        =   25
            ToolTipText     =   "Traz os AWBs com todas as suas respectivas NFs"
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.ComboBox ComboREG 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   3075
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processar / Atualizar Dados"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar pela Região"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   420
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
      TabIndex        =   16
      Top             =   120
      Width           =   5595
      Begin VB.Frame fraPorPeriodo 
         Caption         =   "Período"
         Height          =   1035
         Left            =   1620
         TabIndex        =   20
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
            TabIndex        =   22
            Top             =   240
            Width           =   2160
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1860
            TabIndex        =   21
            Top             =   540
            Width           =   90
         End
      End
      Begin VB.Frame fraPorEmissao 
         Height          =   1035
         Left            =   1620
         TabIndex        =   17
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
         TabIndex        =   19
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
Attribute VB_Name = "frmAIRAcompanha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkProcEntregue_Click()
    If chkProcEntregue.Value = 1 Then
        If optPorEmissao = True And opt60d = True Then
            MsgBox "Para Processar CTC/NF Entregue, Escolha um Período de Até 30 dias ou um Mês Específico !"
            chkProcEntregue.Value = 0
            Exit Sub
        End If
        If optPorPeriodo = True And IsDate(mskPer2) And IsDate(mskPer1) Then
            If (CDate(mskPer2) - CDate(mskPer1)) > 32 Then
                MsgBox "Para Processar CTC/NF Entregue, Escolha um Período de Até 30 dias ou um Mês Específico !"
                chkProcEntregue.Value = 0
                Exit Sub
            End If
        End If
        lblAviso.Visible = False
        If optPorCTC = True Then
            GridEntregueCtc.Visible = True
            lblCtcsEntregue.Visible = True
            cmdGerarArqEntregueCtc.Visible = True
            cmdSACEntregue.Visible = True
        Else
            GridEntregue.Visible = True
            lblNfsEntregue.Visible = True
            cmdGerarArqEntregue.Visible = True
            cmdSACEntregue.Visible = True
        End If
    Else
        lblAviso.Visible = True
        If optPorCTC = True Then
            GridEntregueCtc.Visible = False
            lblCtcsEntregue.Visible = False
            cmdGerarArqEntregueCtc.Visible = False
            cmdSACEntregue.Visible = False
        Else
            GridEntregue.Visible = False
            lblNfsEntregue.Visible = False
            cmdGerarArqEntregue.Visible = False
            cmdSACEntregue.Visible = False
        End If
    End If
End Sub

Private Sub chkProcEntregue_LostFocus()
    If chkProcEntregue.Value = 1 Then
        lblAviso.Visible = False
        GridEntregue.Visible = True
        GridEntregueCtc.Visible = True
    Else
        lblAviso.Visible = True
        GridEntregue.Visible = False
        GridEntregueCtc.Visible = False
    End If
End Sub

Private Sub chkTodosEstab_Click()
    If chkTodosEstab.Value = 1 Then
        txtCGCRem.MaxLength = 8
    Else
        txtCGCRem.MaxLength = 14
    End If
    txtCGCRem.SetFocus
End Sub

Private Sub chkTodosEstab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdBuscaREM_Click()
    frmBuscaCLI.Caption = "Busca Cliente REMETENTE - (Acompanhamento)"
    frmBuscaCLI.Show 1
End Sub

Private Sub cmdGerarArqEntregue_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (ENTREGUE) - POR NF"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqEntregueCtc_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (ENTREGUE) - POR CTC"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqOcorr_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (EM OCORRÊNCIA) - POR NF"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqOcorrCtc_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (EM OCORRÊNCIA) - POR CTC"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqSemPos_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (SEM POSIÇÃO) - POR NF"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqSemPosCtc_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (SEM POSIÇÃO) - POR CTC"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqTransito_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (EM TRÂNSITO) - POR NF"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub cmdGerarArqTransitoCtc_Click()
    frmSelGerarArqAcomp.fraSelecao.Caption = "Seleção dos Dados (EM TRÂNSITO) - POR CTC"
    frmSelGerarArqAcomp.Show 1
End Sub

Private Sub CmdEMAIL_Click()
Dim CIA As String
Dim AWB As String
Dim Dig As String
Dim Peso As String
Dim Volumes As String
Dim Especie As String
Dim DescrProdSis As String
Dim CidadeOrigem As String
Dim UFOrigem As String
Dim SiglaOrigem As String
Dim CidadeDestino As String
Dim UFDestino As String
Dim SiglaDestino As String
Dim Destinatario As String
Dim Email As String
Dim NFs(0 To 500, 0 To 4) As String
Dim Y As Integer
Dim X As Integer
Dim Z As Integer
Dim Texto As String
Dim xCol As Integer
Dim xRow As Integer
Dim xEmailSent As Integer
Dim xEmailRec As Integer

CmdEMAIL.Enabled = False
cmdInserirVOO.Enabled = False
cmdGerarArq.Enabled = False
cmdProcessa.Enabled = False
cmdSair.Enabled = False
Me.MousePointer = 11
DoEvents

Texto = ""
CIA = FlexAWB.TextMatrix(1, 2)
AWB = FlexAWB.TextMatrix(1, 3)
Dig = FlexAWB.TextMatrix(1, 4)
Peso = FlexAWB.TextMatrix(1, 11)
Volumes = FlexAWB.TextMatrix(1, 10)
Especie = FlexAWB.TextMatrix(1, 9)
    If FlexAWB.TextMatrix(1, 17) = "S" Then
    DescrProdSis = FlexAWB.TextMatrix(1, 16) & " - PERECIVEL"
    Else
    DescrProdSis = FlexAWB.TextMatrix(1, 16)
    End If
SiglaOrigem = FlexAWB.TextMatrix(1, 14)
CidadeDestino = FlexAWB.TextMatrix(1, 13)
SiglaDestino = FlexAWB.TextMatrix(1, 15)
Destinatario = FlexAWB.TextMatrix(1, 12)
Email = FlexAWB.TextMatrix(1, 35)

xEmailRec = 0
xEmailSent = 0

X = 0
Y = 1
NFs(X, 0) = FlexAWB.TextMatrix(1, 5)
NFs(X, 1) = FlexAWB.TextMatrix(1, 6)
NFs(X, 2) = FlexAWB.TextMatrix(1, 7)
NFs(X, 3) = FlexAWB.TextMatrix(1, 8)


    If Len(Trim(FlexAWB.TextMatrix(Y, 35))) > 0 Then
    FlexAWB.Row = Y
    xRow = Y
        For xCol = 1 To FlexAWB.Cols - 1
        FlexAWB.Col = xCol
        FlexAWB.CellBackColor = xAmarelo
        Next
    End If

X = X + 1

    For Y = 2 To FlexAWB.Rows - 1
        If CIA = FlexAWB.TextMatrix(Y, 2) And AWB = FlexAWB.TextMatrix(Y, 3) And Dig = FlexAWB.TextMatrix(Y, 4) And Email = FlexAWB.TextMatrix(Y, 35) Then
        NFs(X, 0) = FlexAWB.TextMatrix(Y, 5)
        NFs(X, 1) = FlexAWB.TextMatrix(Y, 6)
        NFs(X, 2) = FlexAWB.TextMatrix(Y, 7)
        NFs(X, 3) = FlexAWB.TextMatrix(Y, 8)
            If Len(Trim(FlexAWB.TextMatrix(Y, 35))) > 0 Then
            FlexAWB.Row = Y
            xRow = Y
                For xCol = 1 To FlexAWB.Cols - 1
                FlexAWB.Col = xCol
                FlexAWB.CellBackColor = xAmarelo
                Next
            End If
        X = X + 1
        Else
        xTexto = xTexto & "AWB: " & CIA & "-" & AWB & "-" & Dig & Chr(13)
        xTexto = xTexto & "Peso: " & Peso & Chr(13)
        xTexto = xTexto & "Volumes: " & Volumes & Chr(13)
        xTexto = xTexto & "Especie: " & Especie & Chr(13)
        xTexto = xTexto & "Descricao: " & DescrProdSis & Chr(13)
        xTexto = xTexto & "Destino: " & CidadeDestino & Chr(13)
        xTexto = xTexto & "Destinatario: " & Destinatario & Chr(13)
        xTexto = xTexto & "Trecho: " & SiglaOrigem & "/" & SiglaDestino & Chr(13)
        xTexto = xTexto & "Notas Fiscais: " & Chr(13)
        
            For Z = 0 To X
            xTexto = xTexto & NFs(Z, 0) & " - "
            xTexto = xTexto & NFs(Z, 1) & " - "
            xTexto = xTexto & NFs(Z, 2) & " - "
            xTexto = xTexto & NFs(Z, 3) & Chr(13)
            Next
        
        xTexto = xTexto & Chr(13)
        xTexto = xTexto & Chr(13)
        
            If Email <> FlexAWB.TextMatrix(Y, 35) Then
            If Len(Trim(Email)) > 0 Then
            Dim xMail As Outlook.Application
            Dim xMensagem As Outlook.MailItem
            Set xMail = CreateObject("outlook.application")
            Set xMensagem = xMail.CreateItem(olMailItem)
            
            xMensagem.To = Email
            xMensagem.Subject = "Relato Automático de Embarques de AWBs"
            xMensagem.Body = xTexto
            xMensagem.Send
            
            xMail.Quit
            
            Set xMensagem = Nothing
            Set xMail = Nothing
            xEmailSent = xEmailSent + 1
            Else
                If Len(Trim(FlexAWB.TextMatrix(Y, 3))) > 0 Then
                xEmailRec = xEmailRec + 1
                End If
            End If
            xTexto = ""
            End If

        
        
            For Z = 0 To X
            NFs(Z, 0) = ""
            NFs(Z, 1) = ""
            NFs(Z, 2) = ""
            NFs(Z, 3) = ""
            NFs(Z, 4) = ""
            Next
        
        CIA = FlexAWB.TextMatrix(Y, 2)
        AWB = FlexAWB.TextMatrix(Y, 3)
        Dig = FlexAWB.TextMatrix(Y, 4)
        Peso = FlexAWB.TextMatrix(Y, 11)
        Volumes = FlexAWB.TextMatrix(Y, 10)
        Especie = FlexAWB.TextMatrix(Y, 9)
            If FlexAWB.TextMatrix(Y, 17) = "S" Then
            DescrProdSis = FlexAWB.TextMatrix(Y, 16) & " - PERECIVEL"
            Else
            DescrProdSis = FlexAWB.TextMatrix(Y, 16)
            End If
        SiglaOrigem = FlexAWB.TextMatrix(Y, 14)
        CidadeDestino = FlexAWB.TextMatrix(Y, 13)
        SiglaDestino = FlexAWB.TextMatrix(Y, 15)
        Destinatario = FlexAWB.TextMatrix(Y, 12)
        Email = FlexAWB.TextMatrix(Y, 35)
        X = 0
        NFs(X, 0) = FlexAWB.TextMatrix(Y, 5)
        NFs(X, 1) = FlexAWB.TextMatrix(Y, 6)
        NFs(X, 2) = FlexAWB.TextMatrix(Y, 7)
        NFs(X, 3) = FlexAWB.TextMatrix(Y, 8)
        
            If Len(Trim(FlexAWB.TextMatrix(Y, 35))) > 0 Then
            FlexAWB.Row = Y
            xRow = Y
                For xCol = 1 To FlexAWB.Cols - 1
                FlexAWB.Col = xCol
                FlexAWB.CellBackColor = xAmarelo
                Next
            End If
        
        X = X + 1
        End If
    Next
    
    If xEmailRec > 0 Then
    MsgBox "ATENÇÃO! Não foram enviados " & xEmailRec & " E-mails por não haver informação de endereço dos mesmos. Por favor, verifique junto a estes representantes e atualize seus cadastros com seus respectivos endereços de E-Mail.", vbExclamation, ""
    End If

    If xEmailSent > 0 Then
    MsgBox "Foram enviados " & xEmailSent & " E-mails com sucesso.", vbInformation, ""
    Else
    MsgBox "Nenhum E-Mail foi enviado.", vbCritical, ""
    End If

CmdEMAIL.Enabled = True
cmdInserirVOO.Enabled = True
cmdGerarArq.Enabled = True
cmdProcessa.Enabled = True
cmdSair.Enabled = True
Me.MousePointer = 0
DoEvents

End Sub

Private Sub cmdGerarArq_Click()
Dim xInf As Boolean
xarquivo = "c:\AWBS EMITIDOS - " & Year(Date) & Month(Date) & Day(Date) & " - " & Hour(Time) & Minute(Time) & Second(Time) & ".TXT"
xInf = False

Open xarquivo For Output As #1

    For Y = 0 To FlexAWB.Rows - 1
    xlinha = ""
        For X = 1 To FlexAWB.Cols - 1
        xlinha = xlinha & FlexAWB.TextMatrix(Y, X) & "#"
        Next
        
    FlexAWB.Row = Y
    FlexAWB.Col = 1
        If FlexAWB.CellBackColor = xAmarelo Then
        xInf = True
        End If
        
        If xInf = False Then
        Print #1, xlinha
        Else
            If ChkIncluir.Value = 1 Then
            Print #1, xlinha
            End If
        End If
    xInf = False
    Next
Close #1

MsgBox "O arquivo texto gerado " & xarquivo & " deve ser aberto no Excel. Ele é delimitado e seu delimitador é '#'.", vbInformation, "Arquivo Gerado com Sucesso!"
End Sub

Private Sub cmdInserirVOO_Click()
xACOMP = True
frmAcompAWB.Show 1
End Sub

Private Sub cmdProcessa_Click()
cmdProcessa.Enabled = False
cmdInserirVOO.Enabled = False
CmdEMAIL.Enabled = False
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

    If de_informa.rsMovimentoAWBNF.State = 1 Then de_informa.rsMovimentoAWBNF.Close
    de_informa.MovimentoAWBNF CDate(DataInicial), CDate(DataFinal), "%", "%", Sigla & "%", "%"
    

    
    If de_informa.rsMovimentoAWBNF.RecordCount = 0 Then
    MsgBox "Sua pesquisa para este período não retornou registro algum.", vbCritical, ""
    cmdProcessa.Enabled = True
    CmdEMAIL.Enabled = True
    cmdGerarArq.Enabled = True
    cmdSair.Enabled = True
    Me.MousePointer = 0
    Exit Sub
    End If

FlexAWB.Clear
    
LblREGRET.Caption = "Registros Retornados: " & de_informa.rsMovimentoAWBNF.RecordCount
FlexAWB.Rows = de_informa.rsMovimentoAWBNF.RecordCount + 2
FlexAWB.Cols = 36
FlexAWB.FixedRows = 1
FlexAWB.FixedCols = 1

    

FlexAWB.TextMatrix(0, 1) = "Filial"
FlexAWB.TextMatrix(0, 2) = "Cia."
FlexAWB.TextMatrix(0, 3) = "AWB"
FlexAWB.TextMatrix(0, 4) = "Dig."
FlexAWB.TextMatrix(0, 5) = "Nota"
FlexAWB.TextMatrix(0, 6) = "Série"
FlexAWB.TextMatrix(0, 7) = "Remetente"
FlexAWB.TextMatrix(0, 8) = "CTC"
FlexAWB.TextMatrix(0, 9) = "Espécie"
FlexAWB.TextMatrix(0, 10) = "Volumes"
FlexAWB.TextMatrix(0, 11) = "Peso Real"
FlexAWB.TextMatrix(0, 12) = "Destinatário"
FlexAWB.TextMatrix(0, 13) = "Destino"
FlexAWB.TextMatrix(0, 14) = "Sigla Orig."
FlexAWB.TextMatrix(0, 15) = "Sigla Dest."
FlexAWB.TextMatrix(0, 16) = "Descrição"
FlexAWB.TextMatrix(0, 17) = "Perecível?"
FlexAWB.TextMatrix(0, 18) = "Vôo"
FlexAWB.TextMatrix(0, 19) = "Data Part."
FlexAWB.TextMatrix(0, 20) = "Hora Part."
FlexAWB.TextMatrix(0, 21) = "Airp. Conexão"
FlexAWB.TextMatrix(0, 22) = "Cidade Conexão"
FlexAWB.TextMatrix(0, 23) = "UF Conexão"
FlexAWB.TextMatrix(0, 24) = "Data Conexão"
FlexAWB.TextMatrix(0, 25) = "Hora Conexão"
FlexAWB.TextMatrix(0, 26) = "Dt. Part. Conexão"
FlexAWB.TextMatrix(0, 27) = "Hora Part. Conexão"
FlexAWB.TextMatrix(0, 28) = "Dt. Chegada"
FlexAWB.TextMatrix(0, 29) = "Hora Chegada"
FlexAWB.TextMatrix(0, 30) = "Retirou?"
FlexAWB.TextMatrix(0, 31) = "Representante"
FlexAWB.TextMatrix(0, 32) = "CNPJ"
FlexAWB.TextMatrix(0, 33) = "Localidade"
FlexAWB.TextMatrix(0, 34) = "UF"
FlexAWB.TextMatrix(0, 35) = "E-Mail"

FlexAWB.ColWidth(0) = 400
FlexAWB.ColWidth(1) = 400
FlexAWB.ColWidth(2) = 400
FlexAWB.ColWidth(3) = 900
FlexAWB.ColWidth(4) = 320
FlexAWB.ColWidth(5) = 900
FlexAWB.ColWidth(6) = 700
FlexAWB.ColWidth(7) = 4000
FlexAWB.ColWidth(8) = 1100
FlexAWB.ColWidth(9) = 1400
FlexAWB.ColWidth(10) = 650
FlexAWB.ColWidth(11) = 650
FlexAWB.ColWidth(12) = 4000
FlexAWB.ColWidth(13) = 3500
FlexAWB.ColWidth(14) = 700
FlexAWB.ColWidth(15) = 700
FlexAWB.ColWidth(16) = 3700
FlexAWB.ColWidth(17) = 2000
FlexAWB.ColWidth(18) = 900
FlexAWB.ColWidth(19) = 1200
FlexAWB.ColWidth(20) = 1200
FlexAWB.ColWidth(21) = 3000
FlexAWB.ColWidth(22) = 3000
FlexAWB.ColWidth(23) = 1050
FlexAWB.ColWidth(24) = 1200
FlexAWB.ColWidth(25) = 1200
FlexAWB.ColWidth(26) = 1200
FlexAWB.ColWidth(27) = 1200
FlexAWB.ColWidth(28) = 1200
FlexAWB.ColWidth(29) = 1200
FlexAWB.ColWidth(30) = 750
FlexAWB.ColWidth(31) = 4000
FlexAWB.ColWidth(32) = 1500
FlexAWB.ColWidth(33) = 3000
FlexAWB.ColWidth(34) = 800
FlexAWB.ColWidth(35) = 5000

X = 0

    Do Until de_informa.rsMovimentoAWBNF.EOF
    X = X + 1
    FlexAWB.TextMatrix(X, 0) = X
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("filial")) = False Then FlexAWB.TextMatrix(X, 1) = de_informa.rsMovimentoAWBNF.Fields("filial")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CIA")) = False Then FlexAWB.TextMatrix(X, 2) = de_informa.rsMovimentoAWBNF.Fields("CIA")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("AWB")) = False Then FlexAWB.TextMatrix(X, 3) = de_informa.rsMovimentoAWBNF.Fields("AWB")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("DIG")) = False Then FlexAWB.TextMatrix(X, 4) = de_informa.rsMovimentoAWBNF.Fields("DIG")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("NOTA")) = False Then FlexAWB.TextMatrix(X, 5) = de_informa.rsMovimentoAWBNF.Fields("NOTA")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("SERIE")) = False Then FlexAWB.TextMatrix(X, 6) = de_informa.rsMovimentoAWBNF.Fields("SERIE")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("REMET_NOME")) = False Then FlexAWB.TextMatrix(X, 7) = de_informa.rsMovimentoAWBNF.Fields("REMET_NOME")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CTC")) = False Then FlexAWB.TextMatrix(X, 8) = de_informa.rsMovimentoAWBNF.Fields("CTC")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("ESPECIE")) = False Then FlexAWB.TextMatrix(X, 9) = de_informa.rsMovimentoAWBNF.Fields("ESPECIE")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("VOLUMES")) = False Then FlexAWB.TextMatrix(X, 10) = de_informa.rsMovimentoAWBNF.Fields("VOLUMES")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("PESOREAL")) = False Then FlexAWB.TextMatrix(X, 11) = de_informa.rsMovimentoAWBNF.Fields("PESOREAL")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("NOMEDES")) = False Then FlexAWB.TextMatrix(X, 12) = de_informa.rsMovimentoAWBNF.Fields("NOMEDES")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CIDADEDES")) = False Then FlexAWB.TextMatrix(X, 13) = de_informa.rsMovimentoAWBNF.Fields("CIDADEDES")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("SIGLAORIGEM")) = False Then FlexAWB.TextMatrix(X, 14) = de_informa.rsMovimentoAWBNF.Fields("SIGLAORIGEM")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("SIGLADES")) = False Then FlexAWB.TextMatrix(X, 15) = de_informa.rsMovimentoAWBNF.Fields("SIGLADES")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("DESCRPRODSIS")) = False Then FlexAWB.TextMatrix(X, 16) = de_informa.rsMovimentoAWBNF.Fields("DESCRPRODSIS")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("PERECIVEL")) = False Then FlexAWB.TextMatrix(X, 17) = de_informa.rsMovimentoAWBNF.Fields("PERECIVEL")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("VOO")) = False Then FlexAWB.TextMatrix(X, 18) = de_informa.rsMovimentoAWBNF.Fields("VOO")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("DATA_PARTIDA")) = False Then FlexAWB.TextMatrix(X, 19) = de_informa.rsMovimentoAWBNF.Fields("DATA_PARTIDA")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("HORA_PARTIDA")) = False Then FlexAWB.TextMatrix(X, 20) = de_informa.rsMovimentoAWBNF.Fields("HORA_PARTIDA")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONAEROPORTO")) = False Then FlexAWB.TextMatrix(X, 21) = de_informa.rsMovimentoAWBNF.Fields("CONAEROPORTO")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONCIDADE")) = False Then FlexAWB.TextMatrix(X, 22) = de_informa.rsMovimentoAWBNF.Fields("CONCIDADE")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONUF")) = False Then FlexAWB.TextMatrix(X, 23) = de_informa.rsMovimentoAWBNF.Fields("CONUF")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONDTCHEG")) = False Then FlexAWB.TextMatrix(X, 24) = de_informa.rsMovimentoAWBNF.Fields("CONDTCHEG")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONHORACHEG")) = False Then FlexAWB.TextMatrix(X, 25) = de_informa.rsMovimentoAWBNF.Fields("CONHORACHEG")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONDTPART")) = False Then FlexAWB.TextMatrix(X, 26) = de_informa.rsMovimentoAWBNF.Fields("CONDTPART")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CONHORAPART")) = False Then FlexAWB.TextMatrix(X, 27) = de_informa.rsMovimentoAWBNF.Fields("CONHORAPART")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("DATA_CHEGADA")) = False Then FlexAWB.TextMatrix(X, 28) = de_informa.rsMovimentoAWBNF.Fields("DATA_CHEGADA")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("HORA_CHEGADA")) = False Then FlexAWB.TextMatrix(X, 29) = de_informa.rsMovimentoAWBNF.Fields("HORA_CHEGADA")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CLIENTERETIROU")) = False Then FlexAWB.TextMatrix(X, 30) = de_informa.rsMovimentoAWBNF.Fields("CLIENTERETIROU")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("NOME")) = False Then FlexAWB.TextMatrix(X, 31) = de_informa.rsMovimentoAWBNF.Fields("NOME")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("CGC")) = False Then FlexAWB.TextMatrix(X, 32) = de_informa.rsMovimentoAWBNF.Fields("CGC")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("LOCALIDADE")) = False Then FlexAWB.TextMatrix(X, 33) = de_informa.rsMovimentoAWBNF.Fields("LOCALIDADE")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("UF")) = False Then FlexAWB.TextMatrix(X, 34) = de_informa.rsMovimentoAWBNF.Fields("UF")
    If IsNull(de_informa.rsMovimentoAWBNF.Fields("EMAIL")) = False Then FlexAWB.TextMatrix(X, 35) = de_informa.rsMovimentoAWBNF.Fields("EMAIL")
    DoEvents
    de_informa.rsMovimentoAWBNF.MoveNext
    Loop
    
cmdProcessa.Enabled = True
CmdEMAIL.Enabled = True
cmdGerarArq.Enabled = True
cmdSair.Enabled = True
Me.MousePointer = 0
DoEvents

End Sub

Private Sub cmdSACBxSemEntr_Click()
    xultimofilial = Mid(GridBaixa.Columns(0), 1, 2)
    xultimoctc = Mid(GridBaixa.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridBaixa.Columns(0), 1, 2)
    frmSac.TxtCTC = Mid(GridBaixa.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACEntregue_Click()
    xultimofilial = Mid(GridEntregueCtc.Columns(0), 1, 2)
    xultimoctc = Mid(GridEntregueCtc.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridEntregueCtc.Columns(0), 1, 2)
    frmSac.TxtCTC = Mid(GridEntregueCtc.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACOcorr_Click()
    xultimofilial = Mid(GridOcorrCTC.Columns(0), 1, 2)
    xultimoctc = Mid(GridOcorrCTC.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridOcorrCTC.Columns(0), 1, 2)
    frmSac.TxtCTC = Mid(GridOcorrCTC.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSACSemPos_Click()
frmAcompAWB.Show 1
End Sub

Private Sub cmdSACTransito_Click()
    xultimofilial = Mid(GridTransitoCTC.Columns(0), 1, 2)
    xultimoctc = Mid(GridTransitoCTC.Columns(0), 3, 8)
    frmSac.TxtFilial = Mid(GridTransitoCTC.Columns(0), 1, 2)
    frmSac.TxtCTC = Mid(GridTransitoCTC.Columns(0), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSelImprSemPos_Click()
    frmSelGerarArqAcomp.Show 1
End Sub


Private Sub FlexAWB_Click()
cmdInserirVOO.Enabled = True
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

End Sub


Private Sub Label2_Click()

End Sub

Private Sub lblNfsEmOcorr_Click()

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


Private Sub optDestinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option4_Click()

End Sub

Private Sub optPer15d_Click()
    mskPer1.Mask = ""
    mskPer1.Text = ""
    mskPer1.Mask = "##/##/####"
    mskPer2.Mask = ""
    mskPer2.Text = ""
    mskPer2.Mask = "##/##/####"
End Sub

Private Sub optPer15d_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPorCTC_Click()
    lblCTCsSemPos.Visible = True
    lblNfsSemPos.Visible = False
    cmdGerarArqSemPosCtc.Visible = True
    cmdGerarArqSemPos.Visible = False
    GridSemPosCtc.Visible = True
    GridSemPos.Visible = False
    cmdSACSemPos.Visible = True
    
    lblCtcsEmOcorr.Visible = True
    lblNfsEmOcorr.Visible = False
    cmdGerarArqOcorrCtc.Visible = True
    cmdGerarArqOcorr.Visible = False
    GridOcorrCTC.Visible = True
    gridOcorr.Visible = False
    cmdSACOcorr.Visible = True
    
    lblCtcsTransito.Visible = True
    lblNfsTransito.Visible = False
    cmdGerarArqTransitoCtc.Visible = True
    cmdGerarArqTransito.Visible = False
    GridTransitoCTC.Visible = True
    GridTransito.Visible = False
    cmdSACTransito.Visible = True
    
    If chkProcEntregue.Value = 1 Then
        lblCtcsEntregue.Visible = True
        lblNfsEntregue.Visible = False
        cmdGerarArqEntregueCtc.Visible = True
        cmdGerarArqEntregue.Visible = False
        GridEntregueCtc.Visible = True
        GridEntregue.Visible = False
        cmdSACEntregue.Visible = True
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

Private Sub optPorNF_Click()
    lblCTCsSemPos.Visible = False
    lblNfsSemPos.Visible = True
    cmdGerarArqSemPosCtc.Visible = False
    cmdGerarArqSemPos.Visible = True
    GridSemPosCtc.Visible = False
    GridSemPos.Visible = True
    cmdSACSemPos.Visible = False
    
    lblCtcsEmOcorr.Visible = False
    lblNfsEmOcorr.Visible = True
    cmdGerarArqOcorrCtc.Visible = False
    cmdGerarArqOcorr.Visible = True
    GridOcorrCTC.Visible = False
    gridOcorr.Visible = True
    cmdSACOcorr.Visible = False
    
    lblCtcsTransito.Visible = False
    lblNfsTransito.Visible = True
    cmdGerarArqTransitoCtc.Visible = False
    cmdGerarArqTransito.Visible = True
    GridTransitoCTC.Visible = False
    GridTransito.Visible = True
    cmdSACTransito.Visible = False
    
    If chkProcEntregue.Value = 1 Then
        lblCtcsEntregue.Visible = False
        lblNfsEntregue.Visible = True
        cmdGerarArqEntregueCtc.Visible = False
        cmdGerarArqEntregue.Visible = True
        GridEntregueCtc.Visible = False
        GridEntregue.Visible = True
        cmdSACEntregue.Visible = False
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

Private Sub optRemetente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optSelCli_Click()
    If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub optSelCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optSelCli_LostFocus()
     If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub optSelReg_Click()
    If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub optSelReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optSelReg_LostFocus()
    If optSelReg.Value = True Then
        fraCliente.Visible = False
        fraRegiao.Visible = True
    Else
        fraCliente.Visible = True
        fraRegiao.Visible = False
    End If
End Sub

Private Sub SSTab5_DblClick()

End Sub

Private Sub txtCGCRem_GotFocus()
    txtCGCRem.SelStart = 0
    txtCGCRem.SelLength = 14
End Sub

Private Sub txtCGCRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub


Private Sub txtregiaosac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub


