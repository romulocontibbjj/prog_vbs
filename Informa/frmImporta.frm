VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImporta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação de Arquivos"
   ClientHeight    =   4785
   ClientLeft      =   2565
   ClientTop       =   1425
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Arquivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   105
      TabIndex        =   10
      Top             =   105
      Width           =   6255
      Begin VB.DirListBox DirImporta 
         Height          =   1890
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtArq 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4200
         MaxLength       =   12
         TabIndex        =   12
         Top             =   1920
         Width           =   1935
      End
      Begin VB.FileListBox fileImport 
         Height          =   1260
         Left            =   4200
         Pattern         =   "*.INF"
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo Escolhido"
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   1680
         Width           =   1275
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
      Height          =   2055
      Left            =   4305
      TabIndex        =   7
      Top             =   2625
      Width           =   2055
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar/Sair"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar Dados"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
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
      Height          =   2055
      Left            =   105
      TabIndex        =   0
      Top             =   2625
      Width           =   4095
      Begin MSComctlLib.ProgressBar progimp 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label label99 
         Height          =   195
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblOcorr 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   20
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Registros Lidos e Gravados de Baixa........:"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   1260
         Width           =   3015
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Registros Lidos e Grav. de Cliente e Dest.:"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   1050
         Width           =   3000
      End
      Begin VB.Label lblcli 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Top             =   1050
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
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblnf 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblctc1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   5
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblctc0 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   3360
         TabIndex        =   4
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Registros Lidos e Gravados de NFs..........:"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registros Lidos e Gravados de CTCs........:"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   630
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registros Lidos de CTCs...........................:"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   420
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim cancela As Boolean   'variável para controle do botão cancelar
    Dim processa As Boolean  'variável para controle do momento de processamento
Private Sub cmdCancel_Click()
    If processa = True Then  'se estiver em processamento de importação de arquivo ???
        cancela = True       'passa esta variável para verdadeiro e aciona opção de cancelamento da importação
    Else
        Unload Me            'caso contrário fecha o form
    End If
End Sub
Private Sub cmdImportar_Click()

'Trava botões do form
    
    cmdCancel.Enabled = False
    cmdImportar.Enabled = False
    DirImporta.Enabled = False
    fileImport.Enabled = False
    txtArq.Enabled = False
    processa = True

'Declara todas as variáveis referente aos campos do arquivo TXT e da tabela.
     
     Dim xfilialctc As String, xfilial As String, xctc As Long, xdata As Date, xhora As String, xremet_cgc As String
     Dim xremet_nome As String, xrespons_cgc As String, xdest_cgc As String, xdest_nome As String
     Dim xcidade_orig As String, xvia As String, xcidade_dest As String, xuf_dest As String, xnfs As String
     Dim xvalmerc As Currency, xpeso_tax As Currency, xpeso_real As Currency, xvolumes As Long, xespecie As String
     Dim xnatureza As String, xdimensoes As String, xfretenacional As Currency, xadvalorem As Currency, xpeso As Currency
     Dim xtxorigem As Currency, xtxdestino As Currency, xtxredespacho As Currency, xtxcoleta As Currency
     Dim xtxoutros As Currency, xfaturar As String, xfaturanum As String, xentrega_data As Date
     Dim xrecebedor As String, xentrega_hora As String, xmodal As String, xobs_emissao As String
     Dim xobs_ocorr As String, xlinha As String, xnumnf As String, xposnf As Long, contctc0 As Long, contctc1 As Long, contnf As Long
     Dim xarq As String, contcli As Long, xtamarq As Long, contocorr As Long, xfretetotal As Currency
     Dim xprazo As Long, xprev_entrega As Date, xfpag As String, xemissor As String, xtranspsub As String
     Dim xendremet As String, xenddest As String, xieremet As String, xiedest As String, xufremet As String, xregiao As String
     Dim xtribut As String, xaliq As Double
     
'Invisivel a barra de progresso e visivel o aguarde até a barra começar a contar
     
     progimp.Visible = False
     lblaguarde.Visible = True
     DoEvents
     
'verifica se o arquivo escolhido para a importação é OK
     
     If txtArq.Text = "" Then
        MsgBox "Erro ! Nome de arquivo Inválido !"
        progimp.Visible = True
        lblaguarde.Visible = False

'destrava os botões no form
        
        cmdCancel.Enabled = True
        cmdImportar.Enabled = True
        DirImporta.Enabled = True
        fileImport.Enabled = True
        txtArq.Enabled = True
        txtArq.SetFocus
        processa = False    'passa processa para falso, processamento interrompido
        Exit Sub
     ElseIf Dir(App.Path & "\CONVERT\" & txtArq.Text) = "" Then
        MsgBox "Erro ! Arquivo Não Encontrado !"
        progimp.Visible = True
        lblaguarde.Visible = False

'destrava os botões no form
        
        cmdCancel.Enabled = True
        cmdImportar.Enabled = True
        DirImporta.Enabled = True
        fileImport.Enabled = True
        txtArq.Enabled = True
        txtArq.SetFocus
        processa = False    'passa processa para falso, processamento interrompido
        Exit Sub
     End If
     
'Abre o arquivo inicialmente para verificar quantos registros tem para o cálculo da Barra de Progresso
     
     xtamarq = 0
     Open App.Path & "\CONVERT\" & txtArq.Text For Input As #1
     Do Until EOF(1)
        Line Input #1, xlinha
        xtamarq = xtamarq + 1
     Loop
     Close #1
     progimp.Max = xtamarq

'Zera os contadores: contctc0=CTC Lidos, contctc1=CTCs Gravador, contnf=Nfs gravadas, contcli=Clientes Gravados
     
     contctc0 = 0
     contctc1 = 0
     contnf = 0
     contcli = 0
     contocorr = 0
     cancela = False   'controle do botão cancelar
     
     lblctc0.Caption = 0
     lblctc1.Caption = 0
     lblnf.Caption = 0
     lblcli.Caption = 0
     lblOcorr.Caption = 0
     
'Abre o arquivo para leitura (txt)
     
     Open App.Path & "\CONVERT\" & txtArq.Text For Input As #1
     
'Inicia o loop para leitura linha a linha do TXT

     cmdCancel.Enabled = True
     progimp.Visible = True
     lblaguarde.Visible = False
     
     
     Do Until EOF(1)
        Line Input #1, xlinha
        If cancela = True Then   'botão cancelar foi pressionado ...
            If MsgBox("Foi Pressionado o botão CANCELAR. Tem certeza que deseja cancelar este processo ?", vbYesNo + vbQuestion, "Interrupção de Processo") = vbYes Then
                frmAtualPrazos.Show 1
                progimp.Visible = True
                progimp.Value = 0
                lblaguarde.Visible = False
                
                'destrava os botões no form
                
                cmdCancel.Enabled = True
                cmdImportar.Enabled = True
                DirImporta.Enabled = True
                fileImport.Enabled = True
                txtArq.Enabled = True
                txtArq.SetFocus
                processa = False    'passa processa para falso, processamento interrompido
                cancela = False
                Close #1
                Exit Sub
                
                'procedure para apagar do arquivo o que já foi incluso deste arquivo
            
            Else
                cancela = False
            End If
        End If

'Se é a primeira linha (concctc0 = 0) verifica se o arquivo já foi processado anteriormente
        
        If contctc0 = 0 Then
            xarq = Mid(xlinha, 1012, 12)
            If de_informa.rsSel_ctcarq.State = 1 Then
               de_informa.rsSel_ctcarq.Close  'se o recordset está aberto fecha-o
            End If
            de_informa.Sel_ctcarq xarq   'busca na tabela o arquivo especificado no TXT
            If de_informa.rsSel_ctcarq.RecordCount > 0 Then
                If MsgBox("ARQUIVO JÁ PROCESSADO ANTERIORMENTE. DESEJA PROCESSÁ-LO NOVAMENTE ? Atenção: Os CTCs que já estiverem cadastrados no sistema não serão cadastrados novamente.", vbYesNo + vbQuestion, "Arquivo Já Processado") = vbNo Then
                    progimp.Visible = True
                    lblaguarde.Visible = False
                    Close #1
            
            'destrava os botões no form
                    
                    cmdCancel.Enabled = True
                    cmdImportar.Enabled = True
                    DirImporta.Enabled = True
                    fileImport.Enabled = True
                    txtArq.Enabled = True
                    processa = False    'passa processa para falso, processamento interrompido
                    Exit Sub
                End If
            End If
        End If
        
'registra nas variáveis os conteúdos dos campos delimitados do TXT

        xfilial = Mid(xlinha, 8, 2)
        xctc = Mid(xlinha, 1, 7)
        xfilialctc = transctc(xfilial, CVar(xctc))
        xdata = CDate(Mid(xlinha, 11, 10))
        xhora = Mid(xlinha, 21, 8)
        xremet_cgc = Mid(xlinha, 30, 14)
        xremet_nome = RTrim(Mid(xlinha, 44, 40))
        xrespons_cgc = Mid(xlinha, 85, 14)
        xdest_cgc = Mid(xlinha, 100, 14)
        xdest_nome = RTrim(Mid(xlinha, 114, 40))
        xcidade_orig = RTrim(Mid(xlinha, 154, 20))
        xvia = RTrim(Mid(xlinha, 174, 3))
        xcidade_dest = RTrim(Mid(xlinha, 177, 20))
        xuf_dest = Mid(xlinha, 197, 2)
        xnfs = Mid(xlinha, 199, 200)
        xvalmerc = Mid(xlinha, 399, 18)
        xvalmerc = xvalmerc / 100
        xpeso_real = Mid(xlinha, 417, 9)
        xpeso_tax = Mid(xlinha, 426, 9)
        If Val(xpeso_real) > Val(xpeso_tax) Then
            xpeso = Val(xpeso_real)
        Else
            xpeso = Val(xpeso_tax)
        End If
        xvolumes = Mid(xlinha, 435, 6)
        xespecie = RTrim(Mid(xlinha, 441, 20))
        xnatureza = RTrim(Mid(xlinha, 461, 20))
        xdimensoes = RTrim(Mid(xlinha, 481, 29))
        xfretenacional = Mid(xlinha, 510, 18)
        xfretenacional = xfretenacional / 100
        xadvalorem = Mid(xlinha, 528, 18)
        xadvalorem = xadvalorem / 100
        xtxorigem = Mid(xlinha, 546, 18)
        xtxorigem = xtxorigem / 100
        xtxdestino = Mid(xlinha, 564, 18)
        xtxdestino = xtxdestino / 100
        xtxredespacho = Mid(xlinha, 582, 18)
        xtxredespacho = xtxredespacho / 100
        xtxcoleta = Mid(xlinha, 600, 18)
        xtxcoleta = xtxcoleta / 100
        xtxoutros = Mid(xlinha, 618, 18)
        xtxoutros = xtxoutros / 100
        xfretetotal = xfretenacional + xadvalorem + xtxorigem + xtxdestino + xtxredespacho + xtxcoleta + xtxoutros
        xfaturar = Mid(xlinha, 636, 1)
        xfaturanum = Mid(xlinha, 637, 10)
        If IsDate(Mid(xlinha, 647, 10)) Then
            xentrega_data = CDate(Mid(xlinha, 647, 10))
        Else
            xentrega_data = CDate("01/01/1900")  'Quando for 01/01/1900
        End If                                   'quer dizer que data é em branco
        xrecebedor = RTrim(Mid(xlinha, 657, 25))
        xentrega_hora = Mid(xlinha, 682, 5)
        xmodal = RTrim(Mid(xlinha, 687, 15))
        xobs_ocorr = RTrim(CStr(Mid(xlinha, 702, 186)))
        xobs_emissao = RTrim(Mid(xlinha, 888, 62) & Mid(xlinha, 950, 62))
        xfpag = RTrim(Mid(xlinha, 1024, 7))
        xemissor = RTrim(Mid(xlinha, 1031, 20))
        xendremet = RTrim(Mid(xlinha, 1051, 40))
        xenddest = RTrim(Mid(xlinha, 1091, 40))
        xieremet = RTrim(Mid(xlinha, 1131, 20))
        xiedest = RTrim(Mid(xlinha, 1151, 20))
        xufremet = Mid(xlinha, 1171, 2)
        xtribut = Mid(xlinha, 1173, 1)
        xaliq = Mid(xlinha, 1174, 7)
        xaliq = xaliq / 10000
        If de_informa.rsSel_ctc_Imp.State = 1 Then
            de_informa.rsSel_ctc_Imp.Close   'fecha o recordset se estiver aberto
        End If
        contctc0 = contctc0 + 1   'atualiza a variável contadora de registros lidos
        lblctc0.Caption = contctc0
        If progimp.Value = 0 Then
           progimp.Visible = True
           lblaguarde.Visible = False
        End If
    
    'Identifica se é Interior / Capital
    
    If de_informa.rsSel_UfCidade.State = 1 Then de_informa.rsSel_UfCidade.Close
    de_informa.Sel_UfCidade Trim$(xuf_dest), Trim$(xcidade_dest)
    If de_informa.rsSel_UfCidade.EOF Then
        xregiao = "INTERIOR"
    Else
        xregiao = "CAPITAL"
    End If
        
    'Tenta Identificar Transportador SubContratada
    
        If de_informa.rsSel_TranspSub.State = 1 Then de_informa.rsSel_TranspSub.Close
        de_informa.Sel_TranspSub
        If de_informa.rsSel_TranspSub.RecordCount = 0 Then
            xtranspsub = "?"
        Else
            Do Until de_informa.rsSel_TranspSub.EOF
                If InStr(1, xobs_emissao, de_informa.rsSel_TranspSub.Fields("texto"), vbTextCompare) > 0 Then
                    xtranspsub = de_informa.rsSel_TranspSub.Fields("transportador")
                    Exit Do
                Else
                    xtranspsub = "?"
                End If
                de_informa.rsSel_TranspSub.MoveNext
            Loop
        End If
            
        
        progimp.Value = contctc0    'atualiza a barra de progresso com base nos registros lidos
      
        
'calcula a previsão de entrega

    If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
    de_informa.Sel_ConsCadCli xremet_cgc
    If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
        xprazo = buscaprazo(xuf_dest, xcidade_dest, de_informa.rsSel_ConsCadCli.Fields("prazo"), Mid$(xmodal, 1, 1))
        xprev_entrega = prev_entr(xdata, xuf_dest, xcidade_dest, xprazo)
    End If

'verifica se o CTC (filial + CTC) já consta na tabela. Caso True não atualiza novamente
        
        de_informa.Sel_ctc_Imp xfilialctc
        If de_informa.rsSel_ctc_Imp.RecordCount = 0 Then
        
'processo que atualiza os dados do cliente do cadastro de cliente (remetentes e destinatários)
            
            'CLIENTE REMETENTE
            If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
            de_informa.Sel_CadCliCGC xremet_cgc
            If de_informa.rsSel_CadCliCGC.RecordCount = 0 Then
                de_informa.ins_cadcli xremet_cgc, xremet_nome, xendremet, xcidade_orig, xufremet, xieremet
                contcli = contcli + 1
                lblcli.Caption = contcli
            Else
                If de_informa.rsSel_CadCliCGC.Fields("nome") <> xremet_nome Or _
                   (de_informa.rsSel_CadCliCGC.Fields("endereco") <> xendremet And Len(Trim$(xendremet)) > 5) Or _
                   de_informa.rsSel_CadCliCGC.Fields("cidade") <> xcidade_orig Or _
                   de_informa.rsSel_CadCliCGC.Fields("uf") <> xufremet Or _
                   de_informa.rsSel_CadCliCGC.Fields("ie") <> xieremet Then
                   de_informa.alt_cadcli_imp xremet_cgc, xremet_nome, xendremet, xcidade_orig, xufremet, xieremet
                   contcli = contcli + 1
                   lblcli.Caption = contcli
                End If
            End If
            'CLIENTE DESTINATÁRIO
            If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
            de_informa.Sel_CadCliCGC xdest_cgc
            If de_informa.rsSel_CadCliCGC.RecordCount = 0 Then
                de_informa.ins_cadcli xdest_cgc, xdest_nome, xenddest, xcidade_dest, xuf_dest, xiedest
                contcli = contcli + 1
                lblcli.Caption = contcli
            Else
                If de_informa.rsSel_CadCliCGC.Fields("nome") <> xdest_nome Or _
                   de_informa.rsSel_CadCliCGC.Fields("endereco") <> xenddest Or _
                   de_informa.rsSel_CadCliCGC.Fields("cidade") <> xcidade_dest Or _
                   de_informa.rsSel_CadCliCGC.Fields("uf") <> xuf_dest Or _
                   de_informa.rsSel_CadCliCGC.Fields("ie") <> xiedest Then
                   de_informa.alt_cadcli_imp xdest_cgc, xdest_nome, xenddest, xcidade_dest, xuf_dest, xiedest
                   contcli = contcli + 1
                   lblcli.Caption = contcli
                End If
            End If
        
'Rotina de inserção na tabela dos dados lidos no CTC

            de_informa.ins_espelhoctc xfilialctc, xfilial, xctc, xdata, xhora, xprev_entrega, xremet_cgc, xremet_nome, xrespons_cgc, _
            xdest_cgc, xdest_nome, xcidade_orig, xvia, xcidade_dest, xuf_dest, xregiao, RTrim(xnfs), xvalmerc, xpeso, _
            xvolumes, xespecie, xnatureza, xdimensoes, xfretenacional, xadvalorem, xtxorigem, _
            xtxdestino, xtxredespacho, xtxcoleta, xtxoutros, xfretetotal, xtribut, xaliq, xfaturar, xfaturanum, xentrega_data, xrecebedor, _
            xentrega_hora, xmodal, xobs_emissao, xobs_ocorr, xarq, xfpag, xemissor, xtranspsub
            
 'se tiver data de entrega, atualiza tb_ocorr com pré-baixa
            
            If xentrega_data <> CDate("01/01/1900") Then 'ou seja, se for diferente de sem data ("01/01/1900" é data em branco)
                de_informa.ins_ocorr1 xfilialctc, xdata, xremet_cgc, "01", "ENTREGA REALIZADA", xentrega_data, xentrega_hora, xentrega_data, xentrega_hora, xrecebedor, "SITLA-LUFT", CVar(Date) & " " & CVar(Time()), "N", Date
                de_informa.alt_temocorr_sn "1", xfilialctc
                contocorr = contocorr + 1
                lblOcorr.Caption = contocorr
                
    'ATENÇÃO: SE FOR UM CTC DA RIACHUELO A BAIXA É AUTOMÁTICA
            ElseIf Mid(xrespons_cgc, 1, 8) = "33200056" Then
                de_informa.ins_ocorr1 xfilialctc, xdata, xremet_cgc, "01", "ENTREGA REALIZADA", xdata + 2, "00:00", xdata + 2, "00:00", "?", "BX.AUTOMÁTICA", CVar(Date) & " " & CVar(Time()), "N", Date
                de_informa.alt_temocorr_sn "1", xfilialctc
                contocorr = contocorr + 1
                lblOcorr.Caption = contocorr
            End If
            
            xnumnf = ""  'variável para separação de cada do campo NFS que está todo junto separado por /
            contctc1 = contctc1 + 1 'Soma a qtde. de CTC gravados
            lblctc1.Caption = contctc1
            
 'processo que separa as NF para gravação no arquivo de NF
    
            For xposnf = 1 To 200   'campo de 200 bytes: string de NFS
                xnumnf = xnumnf & Mid(xnfs, xposnf, 1)
                If (Mid(xnfs, xposnf, 1) = " " And xposnf > 1) Then
                'atualiza o arquivo de NF
                    de_informa.ins_espelhonf RTrim(xnumnf), Val(xnumnf), xfilialctc, xremet_cgc, xremet_nome
                    contnf = contnf + 1  'atualiza a variável de NF gravadas
                    lblnf.Caption = contnf
                    Exit For
                ElseIf Mid(xnfs, xposnf, 2) = "/ " And xposnf > 1 Then
                    xnumnf = Mid(xnumnf, 1, Len(xnumnf) - 1)
                    de_informa.ins_espelhonf RTrim(xnumnf), Val(xnumnf), xfilialctc, xremet_cgc, xremet_nome
                    contnf = contnf + 1
                    lblnf.Caption = contnf
                    Exit For
                ElseIf Mid(xnfs, xposnf, 1) = "/" And xposnf > 1 Then
                    xnumnf = Mid(xnumnf, 1, Len(xnumnf) - 1)
                    de_informa.ins_espelhonf RTrim(xnumnf), Val(xnumnf), xfilialctc, xremet_cgc, xremet_nome
                    contnf = contnf + 1
                    lblnf.Caption = contnf
                    xnumnf = ""
                End If
                DoEvents  'atualiza os eventos pendentes. No caso as qtdes. de registros e a barra de progresso
            Next xposnf
        Else
            'Registro já cadastrado ! Verifica se está com data de entrega e procura
            'ocorrência de Pré baixa, se não tiver pré-baixa , atualiza.
            
            'atualiza a Observação de ocorrência do Sistela SITLA
            de_informa.alt_ObsOcorrSitla RTrim(xobs_ocorr), xfilialctc
            
            'Data de entrega
            If xentrega_data <> CDate("01/01/1900") And de_informa.rsSel_ctc_Imp.Fields("tem_ocorr") <> "1" Then  'ou seja, se for diferente de sem data ("01/01/1900" é data em branco)
                de_informa.ins_ocorr1 xfilialctc, xdata, xremet_cgc, "01", "ENTREGA REALIZADA", xentrega_data, xentrega_hora, xentrega_data, xentrega_hora, xrecebedor, "SITLA-LUFT", CVar(Date) & " " & CVar(Time()), "N", Date
                de_informa.alt_temocorr_sn "1", xfilialctc
                contocorr = contocorr + 1
                lblOcorr.Caption = contocorr
'            ElseIf xentrega_data <> CDate("01/01/1900") And de_informa.rsSel_ctc_Imp.Fields("tem_ocorr") = "1" Then
'                If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
'                de_informa.Sel_ConsOcorr de_informa.rsSel_ctc_Imp.Fields("filialctc"), "01"
'                If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
'                    If xentrega_data <> de_informa.rsSel_ConsOcorr.Fields("data") Then
'                        de_informa.alt_ocorr1 xfilialctc, xentrega_data, xentrega_hora, xentrega_data, xentrega_hora, xrecebedor, "SITLA-LUFT", CVar(Date) & " " & CVar(Time()), "N", Date
'                        contocorr = contocorr + 1
'                        lblOcorr.Caption = contocorr
'                    End If
'                End If
            End If
    
        End If
        DoEvents
    Loop
    Close #1
    frmAtualPrazos.Show 1
    MsgBox "IMPORTAÇÃO / ATUALIZAÇÃO CONCLUÍDA !"

'destrava os botões no form
    
    xtamarq = 0
    cmdCancel.Enabled = True
    cmdImportar.Enabled = True
    DirImporta.Enabled = True
    fileImport.Enabled = True
    txtArq.Enabled = True
    processa = False    'passa processa para falso, processamento interrompido/finalizado
End Sub

Private Sub Command1_Click()
    'If de_informa.rssel_lixo.State = 1 Then de_informa.rssel_lixo.Close
    'de_informa.sel_lixo
    'Do Until de_informa.rssel_lixo.EOF
    '    xnumnf = ""
    '    xnfs = de_informa.rssel_lixo.Fields("nfs") & "                        "
    '    xfilialctc = de_informa.rssel_lixo.Fields("filialctc")
    '    xremet_cgc = de_informa.rssel_lixo.Fields("remet_cgc")
    '    xremet_nome = de_informa.rssel_lixo.Fields("remet_nome")
    '        For xposnf = 1 To 200   'campo de 200 bytes: string de NFS
    '            xnumnf = xnumnf & Mid(xnfs, xposnf, 1)
    '            If (Mid(xnfs, xposnf, 1) = " " And xposnf > 1) Then
                'atualiza o arquivo de NF
    '                de_informa.ins_espelhonf RTrim(xnumnf), xfilialctc, xremet_cgc, xremet_nome
    '                contnf = contnf + 1  'atualiza a variável de NF gravadas
    '                lblnf.Caption = contnf
    '                Exit For
                'ElseIf Mid(xnfs, xposnf, 2) = "/ " And xposnf > 1 Then
                '    xnumnf = Mid(xnumnf, 1, Len(xnumnf) - 1)
                '    de_informa.ins_espelhonf xnumnf, xfilialctc, xremet_cgc, xremet_nome
                '    contnf = contnf + 1
                '    lblnf.Caption = contnf
                '    Exit For
                'ElseIf Mid(xnfs, xposnf, 1) = "/" And xposnf > 1 Then
                '    xnumnf = Mid(xnumnf, 1, Len(xnumnf) - 1)
                '    de_informa.ins_espelhonf xnumnf, xfilialctc, xremet_cgc, xremet_nome
                '    contnf = contnf + 1
                '    lblnf.Caption = contnf
                '    xnumnf = ""
     '           End If
                'DoEvents  'atualiza os eventos pendentes. No caso as qtdes. de registros e a barra de progresso
     '       Next xposnf
     '       de_informa.rssel_lixo.MoveNext
     '   Loop
        
    'MsgBox "fim"
        
End Sub

Private Sub DirImporta_Change()
    fileImport.Path = (DirImporta.Path)  'Quando mudado o diretório, atualiza o path do Dir
    fileImport.Refresh
End Sub
Private Sub fileImport_Click()
    txtArq.Text = fileImport.FileName
End Sub
Private Sub Form_Load()
    cancela = False
    processa = False
    DirImporta.Path = App.Path & "\CONVERT"   'seta o diretório inicial a ser exibido
    frmImporta.Top = 400
    frmImporta.Left = 2500
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuJanela.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If processa = True Then
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuJanela.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmImporta = Nothing
End Sub

Private Sub lixo_Click()
'    Dim xnfchar As String, xnfnum As Long
'    If de_informa.rssel_lixo1.State = 1 Then de_informa.rssel_lixo1.Close
'    de_informa.sel_lixo1
'    Label99.Caption = 0
'    de_informa.rssel_lixo1.MoveFirst
'    Do Until de_informa.rssel_lixo1.EOF
'        Label99.Caption = Val(Label99.Caption) + 1
'        DoEvents
'        xnfchar = de_informa.rssel_lixo1.Fields("numnf")
'        xnfnum = Val(xnfchar)
'        de_informa.alt_lixo1 xnfnum, de_informa.rssel_lixo1.Fields("idcodigo")
'        de_informa.rssel_lixo1.MoveNext
'    Loop
End Sub
