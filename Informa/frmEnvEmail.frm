VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmEnvEmail 
   Caption         =   "Email de Ocorrências para os Clientes"
   ClientHeight    =   6585
   ClientLeft      =   1650
   ClientTop       =   1020
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   7680
   Begin VB.Frame Frame3 
      Caption         =   "Envio de Emails"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   7455
      Begin VB.CheckBox chkOcorrFeriados 
         Caption         =   "Futuros Feriados"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkOcorrCliSac 
         Caption         =   "Ocorrência para Clientes/SAC"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkOcorrIntec 
         Caption         =   "Por Ocorrências (INTEC)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
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
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Processar/Enviar ..."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOG de Emails Enviados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   7455
      Begin MSFlexGridLib.MSFlexGrid flexEmail 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   -1  'True
      Password        =   "olimpio"
      UserName        =   "cassio"
   End
   Begin VB.Label Label3 
      Caption         =   "As mensagens são gravadas na Caixa de Saída e, após enviadas, são armazenadas no Ítens Enviados do OutLook Express."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   $"frmEnvEmail.frx":0000
      Height          =   675
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ATENÇÃO:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmEnvEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEnviar_Click()
    Dim xcgccli As String, xmensagem As String, xdata As Date, xfilialctc As String, xfilialctc2 As String
    Dim xemail1 As String, xemail2 As String, xemail3 As String, xemail4 As String, xemail5 As String, xnome As String
    Dim xdest_nome As String, xcidade_dest As String, xuf_dest As String
    cmdEnviar.Enabled = False
    cmdSair.Enabled = False
    
'tratamento de emails de ocorrências para o cliente
    
If chkOcorrCliSac = 1 Then
    If de_informa.rsSel_OcorrEmailCli.State = 1 Then de_informa.rsSel_OcorrEmailCli.Close
    de_informa.Sel_OcorrEmailCli  'busca as ocorrências a serem enviadas (caso de dúvida verificar o select)
    If de_informa.rsSel_OcorrEmailCli.RecordCount <= 0 Then
        MsgBox "Não há Novos Dados de Ocorrências à serem Informados aos clientes"
        cmdEnviar.Enabled = True
        cmdSair.Enabled = True
    Else
        'rastrea o rs para montar o texto do email
        de_informa.rsSel_OcorrEmailCli.MoveFirst
        Do Until de_informa.rsSel_OcorrEmailCli.EOF
            xcgccli = de_informa.rsSel_OcorrEmailCli.Fields("cgc")
            xdata = de_informa.rsSel_OcorrEmailCli.Fields("data")
            xfilialctc = de_informa.rsSel_OcorrEmailCli.Fields("filialctc")
            xfilialctc2 = Mid(de_informa.rsSel_OcorrEmailCli.Fields("filialctc"), 1, 2) & "-" & Trim(Str(Val((Mid(de_informa.rsSel_OcorrEmailCli.Fields("filialctc"), 3, 8)))))
            xemail1 = Trim(de_informa.rsSel_OcorrEmailCli.Fields("email1"))
            xemail2 = Trim(de_informa.rsSel_OcorrEmailCli.Fields("email2"))
            xemail3 = Trim(de_informa.rsSel_OcorrEmailCli.Fields("email3"))
            xnome = Trim(de_informa.rsSel_OcorrEmailCli.Fields("nome"))
            xmensagem = "Prezado Cliente" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Relacionamos abaixo OCORRÊNCIAS no processo de entrega das seguintes Notas Fiscais:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) + _
            "CLIENTE: " & xnome & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            Do While xcgccli = de_informa.rsSel_OcorrEmailCli.Fields("cgc")
                xdata = de_informa.rsSel_OcorrEmailCli.Fields("data")
                xfilialctc = de_informa.rsSel_OcorrEmailCli.Fields("filialctc")
                xfilialctc2 = Mid(de_informa.rsSel_OcorrEmailCli.Fields("filialctc"), 1, 2) & "-" & Trim(Str(Val((Mid(de_informa.rsSel_OcorrEmailCli.Fields("filialctc"), 3, 8)))))
                xmensagem = xmensagem & _
                            "*******************************************************************" & Chr(13) & Chr(10) & _
                            "Ctc.........:   " & xfilialctc2 & Chr(13) & Chr(10) & _
                            "NF(s).......:   " & Trim(de_informa.rsSel_OcorrEmailCli.Fields("nfs")) & Chr(13) & Chr(10) & _
                            "Destinatário:   " & Trim(de_informa.rsSel_OcorrEmailCli.Fields("dest_nome")) & Chr(13) & Chr(10) & _
                            "Cidade - Uf :   " & Trim(de_informa.rsSel_OcorrEmailCli.Fields("cidade_dest")) & " - " & Trim(de_informa.rsSel_OcorrEmailCli.Fields("uf_dest")) & Chr(13) & Chr(10)
                Do While xfilialctc = Trim(de_informa.rsSel_OcorrEmailCli.Fields("filialctc"))
                    xdata = de_informa.rsSel_OcorrEmailCli.Fields("data")
                    xmensagem = xmensagem & _
                    "--------------------------" & Chr(13) & Chr(10) & _
                            "Data........:   " & xdata & Chr(13) & Chr(10) & _
                            "Ocorrência..:   " & Trim(de_informa.rsSel_OcorrEmailCli.Fields("cod_ocorr")) & " - " & _
                                                 Trim(de_informa.rsSel_OcorrEmailCli.Fields("descr_ocorr")) & Chr(13) & Chr(10)
                    If Len(de_informa.rsSel_OcorrEmailCli.Fields("obs_ocorr")) > 3 Then
                        xmensagem = xmensagem & _
                            "Observação..:   " & Trim(de_informa.rsSel_OcorrEmailCli.Fields("obs_ocorr")) & Chr(13) & Chr(10)
                    End If
                    de_informa.rsSel_OcorrEmailCli.MoveNext
                    If de_informa.rsSel_OcorrEmailCli.EOF Then Exit Do
                Loop
                If de_informa.rsSel_OcorrEmailCli.EOF Then Exit Do
            Loop
            xmensagem = xmensagem & _
                        "*******************************************************************" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Em caso de dúvidas, entrar em contato com nosso SAC no telefone (11) 3602-8900" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Atenciosamente" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "INTEC CARGO" & Chr(13) & Chr(10)
            
            If Len(xemail1) > 4 Then
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail1
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail1
            End If
            If Len(xemail2) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail2
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail2
            End If
            If Len(xemail3) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail3
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail3
            End If
            flexEmail.Rows = Val(flexEmail.Rows) + 1
        Loop
        
        'atualiza a tabela com o LOG de envio (S e Data) Emais enviados para o cliente = "S"
        de_informa.rsSel_OcorrEmailCli.MoveFirst
        Do Until de_informa.rsSel_OcorrEmailCli.EOF
            de_informa.alt_EmailEnvCli datahora("data"), de_informa.rsSel_OcorrEmailCli("codigo")
            de_informa.rsSel_OcorrEmailCli.MoveNext
        Loop
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "PROCESSO", xusuario, "ENV. EMAILS DE INFORMAÇÃO DE OCORRÊNCIAS (PARA OS CLIENTES)"
        
    End If
        
        
'tratar emails para Atendimento aos Cliente (SAC) emails4 e email5 da tabela de cliente

        
    If de_informa.rsSel_OcorrEmailSac.State = 1 Then de_informa.rsSel_OcorrEmailSac.Close
    de_informa.sel_OcorrEmailSac  'busca as ocorrências a serem enviadas (caso de dúvida verificar o select)
    If de_informa.rsSel_OcorrEmailSac.RecordCount <= 0 Then
        MsgBox "Não há Novos Dados de Ocorrências à serem Informados ao SAC (Atendimento)"
        cmdEnviar.Enabled = True
        cmdSair.Enabled = True
    Else
        'rastrea o rs para montar o texto do email
        de_informa.rsSel_OcorrEmailSac.MoveFirst
        Do Until de_informa.rsSel_OcorrEmailSac.EOF
            xcgccli = de_informa.rsSel_OcorrEmailSac.Fields("cgc")
            xdata = de_informa.rsSel_OcorrEmailSac.Fields("data")
            xfilialctc = de_informa.rsSel_OcorrEmailSac.Fields("filialctc")
            xfilialctc2 = Mid(de_informa.rsSel_OcorrEmailSac.Fields("filialctc"), 1, 2) & "-" & Trim(Str(Val((Mid(de_informa.rsSel_OcorrEmailSac.Fields("filialctc"), 3, 8)))))
            xemail4 = Trim(de_informa.rsSel_OcorrEmailSac.Fields("email4"))
            xemail5 = Trim(de_informa.rsSel_OcorrEmailSac.Fields("email5"))
            xnome = Trim(de_informa.rsSel_OcorrEmailSac.Fields("nome"))
            'xdest_nome = Trim(de_informa.rsSel_OcorrEmailSac.Fields("dest_nome"))
            'xcidade_dest = Trim(de_informa.rsSel_OcorrEmailSac.Fields("cidade_dest"))
            'xuf_dest = Trim(de_informa.rsSel_OcorrEmailSac.Fields("uf_dest"))
            xmensagem = "Sr(a). Atendente/SAC" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Abaixo OCORRÊNCIAS no processo de entrega das seguintes Notas Fiscais:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) + _
            "CLIENTE: " & xnome & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            Do While xcgccli = de_informa.rsSel_OcorrEmailSac.Fields("cgc")
                xdata = de_informa.rsSel_OcorrEmailSac.Fields("data")
                xfilialctc = de_informa.rsSel_OcorrEmailSac.Fields("filialctc")
                xfilialctc2 = Mid(de_informa.rsSel_OcorrEmailSac.Fields("filialctc"), 1, 2) & "-" & Trim(Str(Val((Mid(de_informa.rsSel_OcorrEmailSac.Fields("filialctc"), 3, 8)))))
                xmensagem = xmensagem & _
                            "*******************************************************************" & Chr(13) & Chr(10) & _
                            "Ctc.........:   " & xfilialctc2 & Chr(13) & Chr(10) & _
                            "NF(s).......:   " & Trim(de_informa.rsSel_OcorrEmailSac.Fields("nfs")) & Chr(13) & Chr(10) & _
                            "Destinatário:   " & Trim(de_informa.rsSel_OcorrEmailSac.Fields("dest_nome")) & Chr(13) & Chr(10) & _
                            "Cidade - Uf :   " & Trim(de_informa.rsSel_OcorrEmailSac.Fields("cidade_dest")) & " - " & Trim(de_informa.rsSel_OcorrEmailSac.Fields("uf_dest")) & Chr(13) & Chr(10)
                Do While xfilialctc = Trim(de_informa.rsSel_OcorrEmailSac.Fields("filialctc"))
                    xdata = de_informa.rsSel_OcorrEmailSac.Fields("data")
                    xmensagem = xmensagem & _
                            "--------------------------" & Chr(13) & Chr(10) & _
                            "Data........:   " & xdata & Chr(13) & Chr(10) & _
                            "Ocorrência..:   " & Trim(de_informa.rsSel_OcorrEmailSac.Fields("cod_ocorr")) & " - " & _
                                                 Trim(de_informa.rsSel_OcorrEmailSac.Fields("descr_ocorr")) & Chr(13) & Chr(10)
                    If Len(de_informa.rsSel_OcorrEmailSac.Fields("obs_ocorr")) > 3 Then
                        xmensagem = xmensagem & _
                            "Observação..:   " & Trim(de_informa.rsSel_OcorrEmailSac.Fields("obs_ocorr")) & Chr(13) & Chr(10)
                    End If
                    de_informa.rsSel_OcorrEmailSac.MoveNext
                    If de_informa.rsSel_OcorrEmailSac.EOF Then Exit Do
                Loop
                If de_informa.rsSel_OcorrEmailSac.EOF Then Exit Do
            Loop
            xmensagem = xmensagem & _
                        "*******************************************************************" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Em caso de dúvidas, consultar a opção INFORMAÇÃO SAC no Sistema Informa." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Atenciosamente" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Sistema Informa - INTEC TRANSPORTES" & Chr(13) & Chr(10)

            If Len(xemail4) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail4
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail4
            End If
            If Len(xemail5) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail5
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail5
            End If
            flexEmail.Rows = Val(flexEmail.Rows) + 1
        Loop
        
        'atualiza a tabela com o LOG de envio (S e Data) Emais enviados para o cliente = "S"
        de_informa.rsSel_OcorrEmailSac.MoveFirst
        Do Until de_informa.rsSel_OcorrEmailSac.EOF
            de_informa.alt_EmailEnvSac datahora("data"), de_informa.rsSel_OcorrEmailSac("codigo")
            de_informa.rsSel_OcorrEmailSac.MoveNext
        Loop
    
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "PROCESSO", xusuario, "ENV. EMAILS DE INFORMAÇÃO DE OCORRÊNCIAS (EQUIPE ATENDIMENTO - SAC)"
    
    End If
End If


'tratar emails por Ocorrência - Usuários Interno INTEC email1,2,3,4 da tab. de ocorrências


If chkOcorrIntec = 1 Then
    If de_informa.rsSel_OcorrEmailInt.State = 1 Then de_informa.rsSel_OcorrEmailInt.Close
    de_informa.Sel_OcorrEmailInt  'busca as ocorrências a serem enviadas (caso de dúvida verificar o select)
    If de_informa.rsSel_OcorrEmailInt.RecordCount <= 0 Then
        MsgBox "Não há Novos Dados de Ocorrências a Serem Informados aos Usuários Internos INTEC"
        cmdEnviar.Enabled = True
        cmdSair.Enabled = True
    Else
        'rastrea o rs para montar o texto do email
        de_informa.rsSel_OcorrEmailInt.MoveFirst
        Do Until de_informa.rsSel_OcorrEmailInt.EOF
            xcgccli = de_informa.rsSel_OcorrEmailInt.Fields("cgc")
            xCod_Ocorr = de_informa.rsSel_OcorrEmailInt.Fields("cod_ocorr")
            xdata = de_informa.rsSel_OcorrEmailInt.Fields("data")
            xfilialctc = de_informa.rsSel_OcorrEmailInt.Fields("filialctc")
            xfilialctc2 = Mid(de_informa.rsSel_OcorrEmailInt.Fields("filialctc"), 1, 2) & "-" & Trim(Str(Val((Mid(de_informa.rsSel_OcorrEmailInt.Fields("filialctc"), 3, 8)))))
            xemail1 = Trim(de_informa.rsSel_OcorrEmailInt.Fields("email1"))
            xemail2 = Trim(de_informa.rsSel_OcorrEmailInt.Fields("email2"))
            xemail3 = Trim(de_informa.rsSel_OcorrEmailInt.Fields("email3"))
            xemail4 = Trim(de_informa.rsSel_OcorrEmailInt.Fields("email4"))
            xnome = Trim(de_informa.rsSel_OcorrEmailInt.Fields("nome"))
            xmensagem = "Prezado Usuário" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            "Relacionamos abaixo OCORRÊNCIAS no processo de entrega das seguintes Notas Fiscais:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) + _
            "CLIENTE: " & xnome & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            Do While xcgccli = de_informa.rsSel_OcorrEmailInt.Fields("cgc") And xCod_Ocorr = de_informa.rsSel_OcorrEmailInt.Fields("cod_ocorr")
                xdata = de_informa.rsSel_OcorrEmailInt.Fields("data")
                xfilialctc = de_informa.rsSel_OcorrEmailInt.Fields("filialctc")
                xfilialctc2 = Mid(de_informa.rsSel_OcorrEmailInt.Fields("filialctc"), 1, 2) & "-" & Trim(Str(Val((Mid(de_informa.rsSel_OcorrEmailInt.Fields("filialctc"), 3, 8)))))
                xmensagem = xmensagem & _
                            "*******************************************************************" & Chr(13) & Chr(10) & _
                            "Ctc.........:   " & xfilialctc2 & Chr(13) & Chr(10) & _
                            "NF(s).......:   " & Trim(de_informa.rsSel_OcorrEmailInt.Fields("nfs")) & Chr(13) & Chr(10) & _
                            "Destinatário:   " & Trim(de_informa.rsSel_OcorrEmailInt.Fields("dest_nome")) & Chr(13) & Chr(10) & _
                            "Cidade - Uf :   " & Trim(de_informa.rsSel_OcorrEmailInt.Fields("cidade_dest")) & " - " & Trim(de_informa.rsSel_OcorrEmailInt.Fields("uf_dest")) & Chr(13) & Chr(10)
                Do While xfilialctc = Trim(de_informa.rsSel_OcorrEmailInt.Fields("filialctc"))
                    xdata = de_informa.rsSel_OcorrEmailInt.Fields("data")
                    xmensagem = xmensagem & _
                            "--------------------------" & Chr(13) & Chr(10) & _
                            "Data........:   " & xdata & Chr(13) & Chr(10) & _
                            "Ocorrência..:   " & Trim(de_informa.rsSel_OcorrEmailInt.Fields("cod_ocorr")) & " - " & _
                                                 Trim(de_informa.rsSel_OcorrEmailInt.Fields("descr_ocorr")) & Chr(13) & Chr(10)
                    If Len(de_informa.rsSel_OcorrEmailInt.Fields("obs_ocorr")) > 3 Then
                        xmensagem = xmensagem & _
                            "Observação..:   " & Trim(de_informa.rsSel_OcorrEmailInt.Fields("obs_ocorr")) & Chr(13) & Chr(10)
                    End If
                    de_informa.rsSel_OcorrEmailInt.MoveNext
                    If de_informa.rsSel_OcorrEmailInt.EOF Then Exit Do
                Loop
                If de_informa.rsSel_OcorrEmailInt.EOF Then Exit Do
            Loop
            xmensagem = xmensagem & _
                        "*******************************************************************" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Em caso de dúvidas, consultar a opção INFORMAÇÃO SAC no Sistema Informa." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Atenciosamente" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Sistema Informa - INTEC TRANSPORTES" & Chr(13) & Chr(10)

            If Len(xemail1) > 4 Then
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail1
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail1
            End If
            If Len(xemail2) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail2
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail2
            End If
            If Len(xemail3) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail3
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail3
            End If
            If Len(xemail4) > 4 Then
                flexEmail.Rows = Val(flexEmail.Rows) + 1
                MAPISession1.SignOn
                MAPIMessages1.SessionID = MAPISession1.SessionID
                MAPIMessages1.Compose
                MAPIMessages1.RecipAddress = xemail4
                MAPIMessages1.MsgSubject = "Informação de Ocorrência"
                MAPIMessages1.MsgNoteText = xmensagem
                MAPIMessages1.Send False
                MAPISession1.SignOff
                flexEmail.Col = 0
                flexEmail.Row = Val(flexEmail.Rows) - 1
                flexEmail.Text = xnome
                flexEmail.Col = 1
                flexEmail.Text = xemail4
            End If
            
            flexEmail.Rows = Val(flexEmail.Rows) + 1
        Loop
        
        'atualiza a tabela com o LOG de envio (S e Data) Emais enviados para USUARIOS INTEC
        de_informa.rsSel_OcorrEmailInt.MoveFirst
        Do Until de_informa.rsSel_OcorrEmailInt.EOF
            de_informa.alt_EmailEnvInt datahora("data"), de_informa.rsSel_OcorrEmailInt("codigo")
            de_informa.rsSel_OcorrEmailInt.MoveNext
        Loop
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "PROCESSO", xusuario, "ENV. EMAILS DE INFORMAÇÃO DE OCORRÊNCIAS (USUÁRIOS INTERNOS - PROCESSOS)"
        
    End If
End If


'tratar emails para Cliente e Usuários Internos quanto a Futuros Feriados
        
        
If chkOcorrFeriados = 1 Then
    'tratar email de futuros feriados
End If
        
        
        MsgBox "Processo Finalizado ! Certifique que as mensagens saíram da Caixa de Saída, clicando em Enviar/Receber do OutLook Express"
        cmdSair.Enabled = True
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If de_informa.rsSel_Ajuste1.State = 1 Then de_informa.rsSel_Ajuste1.Close
    de_informa.sel_ajuste1
    de_informa.rsSel_Ajuste1.MoveFirst
    lbl1 = 1
    lbl2 = de_informa.rsSel_Ajuste1.RecordCount
    Do Until de_informa.rsSel_Ajuste1.EOF
        If de_informa.rsSel_BuscaSubContratado.State = 1 Then de_informa.rsSel_BuscaSubContratado.Close
        de_informa.Sel_BuscaSubContratado de_informa.rsSel_Ajuste1.Fields("transp_sub")
        If de_informa.rsSel_BuscaSubContratado.RecordCount > 0 Then
            xsub = de_informa.rsSel_BuscaSubContratado.Fields("transportador")
            xperc = de_informa.rsSel_BuscaSubContratado.Fields("percentual")
            xmin = de_informa.rsSel_BuscaSubContratado.Fields("minimo")
            xfretepg = de_informa.rsSel_Ajuste1.Fields("fretetotal") * xperc
            If xfretepg < xmin Then xfretepg = xmin
            de_informa.alt_ajuste1 xfretepg, de_informa.rsSel_Ajuste1.Fields("filialctc")
        Else
            MsgBox "Não Encontrou SubContratado !!!"
        End If
        de_informa.rsSel_Ajuste1.MoveNext
        lbl1 = Val(lbl1) + 1
    DoEvents
    Loop
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    MAPISession1.UserName = "informa"
    MAPISession1.Password = "exodo"
    flexEmail.ColWidth(0) = 4000
    flexEmail.ColWidth(1) = 3050
    flexEmail.Row = 0
    flexEmail.Col = 0
    flexEmail.Text = "CLIENTE"
    flexEmail.Col = 1
    flexEmail.Text = "EMAIL"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmEnvEmail = Nothing
End Sub
