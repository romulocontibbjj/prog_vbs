VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmImportaManif 
   Caption         =   "Importação de Manifestos"
   ClientHeight    =   2025
   ClientLeft      =   3180
   ClientTop       =   2940
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   6660
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status da Importação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4095
      Begin MSComctlLib.ProgressBar progimp 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registros Lidos de Manifesto....................:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2985
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registros Gravados de Manifesto.............:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2985
      End
      Begin VB.Label lblManif0 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblManif1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblaguarde 
         AutoSize        =   -1  'True
         Caption         =   "AGUARDE ..."
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label label99 
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar/Sair"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmImportaManif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancela As Boolean
Private Sub cmdCancel_Click()
    cancela = True
End Sub

Private Sub cmdProcessa_Click()

    On Error GoTo TrataErro2

'Declara todas as variáveis referente aos campos do arquivo TXT e da tabela.
     cmdProcessa.Enabled = False
     
     Dim xfilial As String, xmanif As Long, xfilialmanif As String, xembarcadora As String, xplavaveic As String
     Dim xmotorista As String, xconferente As String, xdatamanif As Variant, xhoramanif As String
     Dim xfilialctc As String, xnummanif As String
     
     
'Invisivel a barra de progresso e visivel o aguarde até a barra começar a contar
     
     progimp.Visible = False
     lblaguarde.Visible = True
     DoEvents
     
'verifica se o arquivo escolhido para a importação é OK
     
     If Dir(App.Path & "\CONVERT\MANIF.TXT") = "" Then
        cmdProcessa.Enabled = True
        Exit Sub
     End If
     
'Abre o arquivo inicialmente para verificar quantos registros tem para o cálculo da Barra de Progresso
     
     xtamarq = 0
     Open App.Path & "\CONVERT\MANIF.TXT" For Input As #1
     Do Until EOF(1)
        Line Input #1, xlinha
        If Mid$(xlinha, 1, 2) = "01" Then
            xtamarq = xtamarq + 1
        End If
     Loop
     Close #1
     progimp.Max = xtamarq

'Zera os contadores: contctc0=CTC Lidos, contctc1=CTCs Gravador, contnf=Nfs gravadas, contcli=Clientes Gravados
     
     xcontmanif0 = 0
     xcontmanif1 = 0
     lblManif0.Caption = 0
     lblManif1.Caption = 0
     
'Abre o arquivo para leitura (txt)
     
     Open App.Path & "\CONVERT\MANIF.TXT" For Input As #1
     
'Inicia o loop para leitura linha a linha do TXT

     progimp.Visible = True
     lblaguarde.Visible = False
     
     'On Error Resume Next
     
     cancela = False
     xnummanif = ""
     
     Do Until EOF(1)
        Line Input #1, xlinha
        
        If cancela = True Then   'botão cancelar foi pressionado ...
            If MsgBox("Foi Pressionado o botão CANCELAR. Tem certeza que deseja cancelar este processo ?", vbYesNo + vbQuestion, "Interrupção de Processo") = vbYes Then
                Close #1
                cmdProcessa.Enabled = True
                Exit Sub
            Else
                cancela = False
            End If
        End If

'registra nas variáveis os conteúdos dos campos delimitados do TXT

        If Mid$(xlinha, 1, 2) = "01" Then
            
            lblManif0 = Val(lblManif0) + 1
            progimp.Value = Val(lblManif0)
            
            xfilial = zeros(Val(Mid$(xlinha, 3, 2)), 2)
            xmanif = Val(Mid$(xlinha, 5, 6))
            xfilialmanif = xfilial & zeros(xmanif, 6)
            xembarcadora = Trim$(Mid$(xlinha, 11, 40))
            xplacaveic = Trim$(Mid$(xlinha, 51, 3) & Mid$(xlinha, 55, 4))
            xmotorista = Trim$(Mid$(xlinha, 59, 30))
            xconferente = Trim$(Mid$(xlinha, 89, 30))
            
            If Val(Mid$(xlinha, 125, 4)) > 2000 Then
                xdatamanif = CDate(Mid$(xlinha, 125, 4) & "/" & zeros(Trim$(Mid$(xlinha, 122, 2)), 2) & "/" & zeros(Trim$(Mid$(xlinha, 119, 2)), 2))
                xhoramanif = zeros(Trim$(Mid$(xlinha, 129, 2)), 2) & ":" & zeros(Trim$(Mid$(xlinha, 132, 2)), 2)
            Else
                xdatamanif = ""
                xhoramanif = ""
            End If
            
        ElseIf Mid$(xlinha, 1, 2) = "02" Then
        
            xfilialctc = zeros(Trim$(Mid$(xlinha, 3, 2)), 2) & zeros(Trim$(Mid$(xlinha, 5, 8)), 8)
            
            If de_informa.rsSel_Manifesto.State = 1 Then de_informa.rsSel_Manifesto.Close
            de_informa.Sel_Manifesto xfilialmanif, xfilialctc
            
            If de_informa.rsSel_Manifesto.RecordCount < 1 Then
                
                If IsDate(xdatamanif) Then
                    If xnummanif <> xfilialmanif Then  'mudou de manifesto
                        lblManif1 = Val(lblManif1) + 1
                    End If
                    xnummanif = xfilialmanif
                    de_informa.Ins_Manifesto xfilial, xmanif, xfilialmanif, xfilialctc, xembarcadora, xplacaveic, _
                    xmotorista, xconferente, xdatamanif, xhoramanif
                End If
            
            End If
        
        End If
        DoEvents
     Loop
     
    Close #1
    cmdProcessa.Enabled = True
    
    Exit Sub
    
TrataErro1:

    de_informa.cn_informa.RollbackTrans
    frmErro.lblErro = "(" & Err.Number & ") " & Err.Description
    frmErro.lblMensagemCompl = ""
    frmErro.Show 1
    Unload Me
    End
    
TrataErro2:

    frmErro.lblErro = "(" & Err.Number & ") " & Err.Description
    frmErro.lblMensagemCompl = ""
    frmErro.Show 1
    Unload Me
    End
    
     
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    cmdProcessa_Click
    Me.Hide
End Sub
