VERSION 5.00
Begin VB.Form frmImprEtiquetas 
   Caption         =   "Impressão de Etiquetas"
   ClientHeight    =   3045
   ClientLeft      =   1500
   ClientTop       =   1305
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      Caption         =   "Etiquetas de Endereçamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtFat2 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFat1 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFilia 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   0
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Fatura de:"
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmImprEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    If de_informa.rsSel_FaturaInterval.State = 1 Then de_informa.rsSel_FaturaInterval.Close
    de_informa.Sel_FaturaInterval TransFatur(Trim$(txtFilia), Trim$(txtFat1)), TransFatur(Trim$(txtFilia), Trim$(txtFat2))
    
    If de_informa.rsSel_FaturaInterval.RecordCount > 0 Then
    
        'busca impressora para este documento
        If Dir("C:\informa.cfg") <> "" Then
            
            Open "C:\informa.cfg" For Input As #1
            
            Do Until EOF(1)
                Line Input #1, xlinha
                If Mid$(xlinha, 1, 3) = "FAT" Then
                    If Trim$(Mid$(xlinha, 5, 2)) = "\\" Or Trim$(Mid$(xlinha, 5, 2)) = "**" Then
                        ximpr_cfg = Trim$(Mid$(xlinha, 5))
                    Else
                        ximpr_cfg = "LPT1"
                    End If
                    Exit Do
                End If
            Loop
            
            If EOF(1) Then
                MsgBox "Não está Configurado a Impressora para Este Documento: ETIQUETA "
                Close #1
                Exit Sub
            End If
            
            Close #1
            
        Else
        
            MsgBox "Não está Configurado a Impressora para Este Documento."
            Exit Sub
            
        End If

        'seta impressora
        
        If ximpr_cfg = "LPT1" Then
            Open ximpr_cfg For Output As #1
            DoEvents
        Else
            For Each ximpr_inst In Printers
                If ximpr_inst.DeviceName = ximpr_cfg Then
                    Open ximpr_cfg For Output As #1
                    DoEvents
                    Exit For
                End If
            Next
        End If

        Do Until de_informa.rsSel_FaturaInterval.EOF

            'inicia a impressão
            Print #1, "Nro. Fatura: " & Mid$(de_informa.rsSel_FaturaInterval.Fields("filialfatura"), 1, 2) & "-" & Mid$(de_informa.rsSel_FaturaInterval.Fields("filialfatura"), 3)
            Print #1, ""
            Print #1, de_informa.rsSel_FaturaInterval.Fields("cliente_nome")
            Print #1, de_informa.rsSel_FaturaInterval.Fields("endcob")
            Print #1, Trim$(de_informa.rsSel_FaturaInterval.Fields("cidadecob")) & "-" & de_informa.rsSel_FaturaInterval.Fields("ufcob")
            Print #1, Mid$(de_informa.rsSel_FaturaInterval.Fields("cepcob"), 1, 5) & "-" & Mid$(de_informa.rsSel_FaturaInterval.Fields("cepcob"), 6, 3)
            Print #1, ""
            If Len(Trim$(de_informa.rsSel_FaturaInterval.Fields("contatocob"))) > 0 Then
                Print #1, "A/C: "; de_informa.rsSel_FaturaInterval.Fields("contatocob")
            Else
                Print #1, "A/C: CONTAS A PAGAR"
            End If
            Print #1, ""
    
            de_informa.rsSel_FaturaInterval.MoveNext
            
        Loop
        
        Close #1
    
    Else
    
        MsgBox "Não Há Dados Neste Intervalo de Faturas !"
        txtFilia.SetFocus
    
    End If
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
