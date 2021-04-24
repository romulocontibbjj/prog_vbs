VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAtualPrazos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualizando ..."
   ClientHeight    =   1365
   ClientLeft      =   3480
   ClientTop       =   2460
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar progresso 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblFilialctc 
      AutoSize        =   -1  'True
      Caption         =   "%"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Atualizando Prazo de Entrega. Aguarde ..."
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "frmAtualPrazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Set frmAtualPrazos = Nothing
End Sub
Private Sub Timer1_Timer()
    Dim xcodigo As Long, xprogresso As Long, xmodal As String, xuf As String, xCidade As String, xtab As String
    Dim xbuscaprazo As String, xdiasprazo As Long, xabonodias As String, xprazo_TT As Integer, xrs As Recordset
    
    Timer1.Interval = 0
    DoEvents
    On Error Resume Next
    mdiInforma.StatusBar1.Panels.Item(2).Text = "Aguarde ... Processandos Prazos de Entrega"
    
    If lblFilialctc = "%" Then
        Set xrs = de_informa.rsSel_AtualPrazo
        If de_informa.rsSel_AtualPrazo.State = 1 Then de_informa.rsSel_AtualPrazo.Close
        de_informa.Sel_AtualPrazo
    Else
        Set xrs = de_informa.rsSel_AtualPrazoCTC
        If de_informa.rsSel_AtualPrazoCTC.State = 1 Then de_informa.rsSel_AtualPrazoCTC.Close
        de_informa.Sel_AtualPrazoCTC lblFilialctc
    End If
    
    If xrs.RecordCount > 0 Then
        xrs.MoveFirst
        DoEvents
        progresso.Max = xrs.RecordCount
        progresso.Value = 0
        Do Until xrs.EOF
        
            xprogresso = xprogresso + 1
            progresso.Value = xprogresso
            DoEvents
            xcodigo = xrs.Fields("codigo")
            xmodal = Mid(xrs.Fields("modal"), 1, 1)
            xuf = xrs.Fields("uf_dest")
            xCidade = xrs.Fields("cidade_dest")
            If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
            de_informa.Sel_ConsCadCli xrs.Fields("remet_cgc")
            If de_informa.rsSel_ConsCadCli.RecordCount = 0 Then
                MsgBox "CGC do cliente não encontrado. Erro de Consistência. Chame Suporte Técnico"
                Unload Me
                Exit Sub
            Else
                If Mid$(de_informa.rsSel_ConsCadCli.Fields("prazo"), 1, 3) <> "TAB" Then
                    xtab = "TAB000"
                Else
                    xtab = de_informa.rsSel_ConsCadCli.Fields("prazo")
                End If
            End If
            
            If xrs.Fields("motivodoc") = "DEV" Then
                If xrs.Fields("prev_entregatipo") = "I" Then
                    xprazo_TT = diasprazo(xrs.Fields("emissaoctc"), xrs.Fields("prev_entrega"), _
                                   xrs.Fields("remet_uf"), xrs.Fields("remet_cidade"), _
                                   xrs.Fields("hora"), xrs.Fields("modal"), _
                                   Mid$(xrs.Fields("filialctc"), 1, 2))
                Else
                    xbuscaprazo = buscaprazo2(xuf, xCidade, "TAB000", xmodal) 'DEVOLUÇÃO TAB PRAZO TAB000
                    xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
                    
                    'verifica horário de corte - HORA
                    If Val(Mid$(xrs.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                        'emissao posterior ao horário de corte
                        xprazo_TT = xprazo_TT + 1
                    ElseIf Val(Mid$(xrs.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                           Val(Mid$(xrs.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                        'emissao posterior ao horário de corte
                        xprazo_TT = xprazo_TT + 1
                    Else
                        If diautil(xrs.Fields("emissaoctc"), xrs.Fields("remet_uf"), _
                           xrs.Fields("remet_cidade")) = False And xprazo_TT = 0 Then
                            xprazo_TT = xprazo_TT + 1
                        End If
                    End If
                End If
            Else
                If xrs.Fields("prev_entregatipo") = "I" Then
                    xprazo_TT = diasprazo(xrs.Fields("emissaoctc"), xrs.Fields("prev_entrega"), _
                                   xrs.Fields("uf_dest"), xrs.Fields("cidade_dest"), _
                                   xrs.Fields("hora"), xrs.Fields("modal"), _
                                   Mid$(xrs.Fields("filialctc"), 1, 2))
                Else
                    xbuscaprazo = buscaprazo2(xuf, xCidade, xtab, xmodal)
                    xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
                    
                    'verifica horário de corte - HORA
                    If Val(Mid$(xrs.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                        'emissao posterior ao horário de corte
                        xprazo_TT = xprazo_TT + 1
                    ElseIf Val(Mid$(xrs.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                           Val(Mid$(xrs.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                        'emissao posterior ao horário de corte
                        xprazo_TT = xprazo_TT + 1
                    Else
                        If diautil(xrs.Fields("emissaoctc"), xrs.Fields("uf_dest"), _
                           xrs.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                            xprazo_TT = xprazo_TT + 1
                        End If
                    End If
                End If
            End If
            
            xdiasprazo = diasprazo(xrs.Fields("emissaoctc"), _
                                   xrs.Fields("data"), _
                                   xuf, xCidade, xrs.Fields("hora"), xmodal, _
                                   Mid$(xrs.Fields("filialctc"), 1, 2))
                                   
            de_informa.alt_diasprazo xcodigo, xprazo_TT, xdiasprazo
            
            'abonos automáticos para os casos de atrasos
            If xdiasprazo > xprazo_TT Then
                'para os casos de prev_entrega informado sendo diferente do TT
                If xrs.Fields("motivodoc") = "DEV" Or _
                   xrs.Fields("motivodoc") = "TRA" Then
                     xabonodias = xdiasprazo - xprazo_TT
                     de_informa.Alt_AbonoAtraso xabonodias, "AUTOMATIC", datahora("DATAHORA"), "DEVOLUÇÃO/TRANSFERÊCIA", xrs.Fields("filialctc")
                Else
                    If de_informa.rsSel_CTCOcorr26e85.State = 1 Then de_informa.rsSel_CTCOcorr26e85.Close
                    'verificar se este CTC possui a Ocorr 26 ou 85
                    de_informa.Sel_CTCOcorr26e85 xrs.Fields("filialctc")
                    If de_informa.rsSel_CTCOcorr26e85.RecordCount > 0 Then  'possui esta ocorrencia
                         'incluir o abono
                         xabonodias = xdiasprazo - xprazo_TT
                         de_informa.Alt_AbonoAtraso xabonodias, "AUTOMATIC", datahora("DATAHORA"), "DEVIDO OCORRÊNCIA", xrs.Fields("filialctc")
                    End If
                End If
            End If
            
            xrs.MoveNext
            
        Loop
    End If
    mdiInforma.StatusBar1.Panels.Item(2).Text = ""
    Unload Me

End Sub
