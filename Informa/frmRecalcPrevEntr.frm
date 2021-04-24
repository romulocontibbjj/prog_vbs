VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRecalcPrevEntr 
   Caption         =   "Recalcular Previsão de Entrega"
   ClientHeight    =   2640
   ClientLeft      =   2610
   ClientTop       =   1785
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   6270
   Begin VB.Frame fraDados 
      Caption         =   "Seleção do Período (data de emissão)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6045
      Begin VB.TextBox txtCgc 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "%"
         Top             =   1680
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar progress1 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa..."
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2625
         TabIndex        =   2
         Top             =   420
         Width           =   1170
         _ExtentX        =   2064
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
         Left            =   1155
         TabIndex        =   1
         Top             =   420
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período:  De"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2415
         TabIndex        =   6
         Top             =   420
         Width           =   90
      End
      Begin VB.Label lblaguarde 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmRecalcPrevEntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcessa_Click()
    Dim xprazointec As String, xprazo_TT As Integer
    
    If Not IsDate(mskPer1) Or Not IsDate(mskPer2) Then
        MsgBox "Período Inválido !"
        mskPer1.SetFocus
        Exit Sub
    End If
    If CDate(mskPer1) > CDate(mskPer2) Then
        MsgBox "Período Inválido !"
        mskPer1.SetFocus
        Exit Sub
    End If
    
    lblaguarde.Caption = "Aguarde Processamento..."
    DoEvents
    If de_informa.rsSel_CtcsPeriodo.State = 1 Then de_informa.rsSel_CtcsPeriodo.Close
    de_informa.Sel_CtcsPeriodo mskPer1, mskPer2, txtCgc
    progress1.Visible = True
    progress1.Max = de_informa.rsSel_CtcsPeriodo.RecordCount
    progresso = 0
    DoEvents
    Do Until de_informa.rsSel_CtcsPeriodo.EOF
    
        progresso = progresso + 1
        progress1.Value = progresso
        DoEvents
        
'        If de_informa.rsSel_CtcsPeriodo.Fields("filialctc") = "0310200482" Then
'            MsgBox "ok"
'        End If
        
        If de_informa.rsSel_CtcsPeriodo.Fields("motivodoc") = "DEV" Then
            
            xprazointec = buscaprazo2(de_informa.rsSel_CtcsPeriodo.Fields("remet_uf"), _
                               Trim$(de_informa.rsSel_CtcsPeriodo.Fields("remet_cidade")), _
                                     "TAB000", _
                                     Mid$(de_informa.rsSel_CtcsPeriodo.Fields("modal"), 1, 1))
                                     
            xprazo_TT = Val(Mid$(xprazointec, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rsSel_CtcsPeriodo.Fields("hora"), 1, 2)) > Val(Mid$(xprazointec, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rsSel_CtcsPeriodo.Fields("hora"), 1, 2)) = Val(Mid$(xprazointec, 4, 2)) And _
                   Val(Mid$(de_informa.rsSel_CtcsPeriodo.Fields("hora"), 4, 2)) > Val(Mid$(xprazointec, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rsSel_CtcsPeriodo.Fields("data"), de_informa.rsSel_CtcsPeriodo.Fields("remet_uf"), _
                   de_informa.rsSel_CtcsPeriodo.Fields("remet_cidade")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
                                     
            xpreventr = prev_entr(CDate(de_informa.rsSel_CtcsPeriodo.Fields("data")), _
                                        de_informa.rsSel_CtcsPeriodo.Fields("remet_uf"), _
                                  Trim$(de_informa.rsSel_CtcsPeriodo.Fields("remet_cidade")), _
                                  xprazo_TT, de_informa.rsSel_CtcsPeriodo.Fields("modal"), _
                                  de_informa.rsSel_CtcsPeriodo.Fields("hora"), _
                                  Mid$(de_informa.rsSel_CtcsPeriodo.Fields("filialctc"), 1, 2))
                                     
        Else
            
            If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
            de_informa.Sel_CadCliCGC de_informa.rsSel_CtcsPeriodo.Fields("remet_cgc")
            xprazointec = buscaprazo2(de_informa.rsSel_CtcsPeriodo.Fields("uf_dest"), _
                               Trim$(de_informa.rsSel_CtcsPeriodo.Fields("cidade_dest")), _
                                     de_informa.rsSel_CadCliCGC.Fields("prazo"), _
                                     de_informa.rsSel_CtcsPeriodo.Fields("modal"))
        
            xprazo_TT = Val(Mid$(xprazointec, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rsSel_CtcsPeriodo.Fields("hora"), 1, 2)) > Val(Mid$(xprazointec, 4, 2)) Then
                'emissao posterior ao horário de corte (hora)
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rsSel_CtcsPeriodo.Fields("hora"), 1, 2)) = Val(Mid$(xprazointec, 4, 2)) And _
                   Val(Mid$(de_informa.rsSel_CtcsPeriodo.Fields("hora"), 4, 2)) > Val(Mid$(xprazointec, 7, 2)) Then
                'emissao posterior ao horário de corte (hora e minuto)
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rsSel_CtcsPeriodo.Fields("data"), de_informa.rsSel_CtcsPeriodo.Fields("uf_dest"), _
                   de_informa.rsSel_CtcsPeriodo.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            
            xpreventr = prev_entr(CDate(de_informa.rsSel_CtcsPeriodo.Fields("data")), _
                                        de_informa.rsSel_CtcsPeriodo.Fields("uf_dest"), _
                                  Trim$(de_informa.rsSel_CtcsPeriodo.Fields("cidade_dest")), _
                                  xprazo_TT, de_informa.rsSel_CtcsPeriodo.Fields("modal"), _
                                  de_informa.rsSel_CtcsPeriodo.Fields("hora"), _
                                  Mid$(de_informa.rsSel_CtcsPeriodo.Fields("filialctc"), 1, 2))
        
        End If
    
        de_informa.alt_PrevEntrega CDate(xpreventr), de_informa.rsSel_CtcsPeriodo.Fields("filialctc")
        de_informa.rsSel_CtcsPeriodo.MoveNext
    
    Loop
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "RECALCULAR PREVISÃO DE ENTREGA"
    
    lblaguarde.Caption = "OK. Processo Finalizado !"
    cmdSair.SetFocus
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRecalcPrevEntr = Nothing
End Sub

Private Sub mskPer1_GotFocus()
    mskPer1.SelStart = 0
    mskPer1.SelLength = 10
End Sub

Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer1_LostFocus()
    mskPer1.Text = century(mskPer1.Text)
End Sub

Private Sub mskPer2_GotFocus()
    mskPer2.SelStart = 0
    mskPer2.SelLength = 10
End Sub

Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer2_LostFocus()
    mskPer2.Text = century(mskPer2.Text)
End Sub

