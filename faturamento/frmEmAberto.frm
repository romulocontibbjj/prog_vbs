VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEmAberto 
   Caption         =   "Títulos em Aberto (À Vencer e Vencidos)"
   ClientHeight    =   6630
   ClientLeft      =   1155
   ClientTop       =   1020
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   7695
   Begin VB.Frame Frame2 
      Caption         =   "Títulos À Vencer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   7455
      Begin VB.CommandButton cmdGeraArq 
         Caption         =   "Gera TXT ..."
         Height          =   350
         Left            =   5280
         TabIndex        =   7
         Top             =   4905
         Width           =   1815
      End
      Begin VB.TextBox lblTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   6
         Top             =   4920
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexDadosResumo 
         Height          =   4455
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblLabelTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total à Vencer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   4920
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optVencidos 
         Caption         =   "Vencidos"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optAvencer 
         Caption         =   "À Vencer"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmEmAberto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirResumo_Click()

End Sub

Private Sub cmdGeraArq_Click()
    Dim xcont As Long, xFiles As String, xlinha As String
    If flexDadosResumo.Rows < 3 Then
        MsgBox "Não há Dados Para Geração de Arquivos !", vbExclamation, "Ops"
    Else
        
        xFiles = "C:\INFORMA\" & "A_VENCER_" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & ".txt"
        
        Open xFiles For Output As #1
    
        Print #1, "FATURAS A VENCER#"
        Print #1, "#"
        Print #1, "SEMANA#VALOR#"
        
        For xcont = 0 To flexDadosResumo.Rows - 1
            xlinha = flexDadosResumo.TextMatrix(xcont, 1) & "#" & _
                     flexDadosResumo.TextMatrix(xcont, 2) & "#"
            Print #1, xlinha
            DoEvents
        Next
        
        Print #1, "#"
        Print #1, lblLabelTotal & "#" & SoNumeros(lblTotal) / 100
        Print #1, "#"
        
        Close #1
        
        MsgBox "OK ! Arquivo Gerado em " & xFiles & "." & Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
        "O Arquivo Gerado é do Tipo Texto ( TXT com Delimitador # ) e você pode abrí-lo em diversos aplicativos. Para Abrí-lo no MS-Excel, em ABRIR escolha ARQUIVOS DO TIPO = Arquivos de Texto e selecione o arquivo no local indicado acima. Na Caixa ASSISTENTE DE IMPORTAÇÃO escolha DELIMITADO e o caracter delimitador escolha OUTROS e digite # . Clique em Concluir e o arquivo será importado para o MS-Excel.", vbInformation, "Geração de Arquivo TXT"

    End If

End Sub

Private Sub cmdProcessar_Click()
    Dim xDataMin As Date, xDataMax As Date, xData1 As Date, xContData As Date
    Dim xvalortot As Currency, xlin As Integer, xSemana As String

    flexDadosResumo.Rows = 1
    flexDadosResumo.Rows = 2
    flexDadosResumo.FixedRows = 1

    If de_informa.rsSel_FatEmAbertoAVencer.State = 1 Then de_informa.rsSel_FatEmAbertoAVencer.Close
    de_informa.Sel_FatEmAbertoAVencer
    
    If de_informa.rsSel_FatEmAbertoAVencer.RecordCount < 1 Then
        MsgBox "Não Há Fatura em Aberto À Vencer !!", vbInformation
        Exit Sub
    Else
        de_informa.rsSel_FatEmAbertoAVencer.MoveFirst
        xDataMin = de_informa.rsSel_FatEmAbertoAVencer.Fields("vencimento")
        de_informa.rsSel_FatEmAbertoAVencer.MoveLast
        xDataMax = de_informa.rsSel_FatEmAbertoAVencer.Fields("vencimento")
        
        xData1 = xDataMin
        xvalortot = 0
        xlin = 1
        
        For xContData = xDataMin To xDataMax Step 1
            
            de_informa.rsSel_FatEmAbertoAVencer.MoveFirst
            xtotalgeral = 0
            Do Until de_informa.rsSel_FatEmAbertoAVencer.EOF
                xtotalgeral = xtotalgeral + de_informa.rsSel_FatEmAbertoAVencer.Fields("valor")
                If xContData = de_informa.rsSel_FatEmAbertoAVencer.Fields("vencimento") Then
                    xvalortot = xvalortot + de_informa.rsSel_FatEmAbertoAVencer.Fields("valor")
                End If
                de_informa.rsSel_FatEmAbertoAVencer.MoveNext
            Loop
            
            If Weekday(xContData) = 1 Or xContData = xDataMax Then
                'incluir valor na flex
                xSemana = xData1 & " à " & xContData
                flexDadosResumo.TextMatrix(xlin, 1) = xSemana
                flexDadosResumo.TextMatrix(xlin, 2) = Format(xvalortot, "##,###,##0.00")
                xvalortot = 0
                xData1 = xContData + 1
                xlin = xlin + 1
                flexDadosResumo.Rows = xlin + 1
            End If
            
        Next
        
    End If
    
    lblTotal = Format(xtotalgeral, "##,###,##0.00")

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If de_informa.rsSel_FatEmAbertoAVencer.State = 1 Then de_informa.rsSel_FatEmAbertoAVencer.Close
    
    flexDadosResumo.Rows = 2
    flexDadosResumo.Cols = 3
    flexDadosResumo.Row = 0
    flexDadosResumo.Col = 1
    flexDadosResumo.Text = "Semana"
    flexDadosResumo.Col = 2
    flexDadosResumo.Text = "Valor"
    
    flexDadosResumo.ColAlignment(1) = 1
    
    flexDadosResumo.ColWidth(0) = 400
    flexDadosResumo.ColWidth(1) = 4000
    flexDadosResumo.ColWidth(2) = 2000
    
End Sub
