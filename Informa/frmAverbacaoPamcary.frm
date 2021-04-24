VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAverbacaoPamcary 
   Caption         =   "Gera Arquivo de Averbação (Pamcary)"
   ClientHeight    =   1995
   ClientLeft      =   2430
   ClientTop       =   2055
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   7230
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar..."
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtCGCCli 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   360
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "54516661%"
      Top             =   720
      Width           =   1590
   End
   Begin VB.Frame Frame6 
      Caption         =   "No Período de ..."
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   360
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CGC do Cliente/Remetente:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1980
   End
End
Attribute VB_Name = "frmAverbacaoPamcary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub cmdGerar_Click()
    Dim xcontador As Long
    
    If de_informa.rsSel_NFsPancary.State Then de_informa.rsSel_NFsPancary.Close
    de_informa.Sel_NFsPancary mskPer1, mskPer2, txtCGCCli
    
    If de_informa.rsSel_NFsPancary.RecordCount < 1 Then
        MsgBox "Não Há Dados para Esta Seleção !"
        Exit Sub
    Else
        Open "C:\INFORMA\PAMCARY\XXXMMDDS.TXT" For Output As #1
    End If
    
    xcontador = 0
    
    Do Until de_informa.rsSel_NFsPancary.EOF
        
    
        'trata data de embarque
        
        xdata = zeros(Day(de_informa.rsSel_NFsPancary.Fields("data")), 2) & _
                zeros(Month(de_informa.rsSel_NFsPancary.Fields("data")), 2) & _
                Mid$(Year(de_informa.rsSel_NFsPancary.Fields("data")), 3, 2)
                
        If de_informa.rsSel_MotoristaPamcary.State = 1 Then de_informa.rsSel_MotoristaPamcary.Close
        de_informa.Sel_MotoristaPamcary Trim$(de_informa.rsSel_NFsPancary.Fields("motorista")) & "%"
        If de_informa.rsSel_MotoristaPamcary.RecordCount < 1 Then
            If de_informa.rsSel_MotoristaPamcary.State = 1 Then de_informa.rsSel_MotoristaPamcary.Close
            de_informa.Sel_MotoristaPamcary "%" & Trim$(de_informa.rsSel_NFsPancary.Fields("motorista")) & "%"
            If de_informa.rsSel_MotoristaPamcary.RecordCount < 1 Then
                If de_informa.rsSel_MotoristaPamcary.State = 1 Then de_informa.rsSel_MotoristaPamcary.Close
                de_informa.Sel_MotoristaPamcary "J%"
            End If
        End If
        If Len(Trim$(de_informa.rsSel_NFsPancary.Fields("placaveic"))) < 7 Then
            xplacaveic = String(7 - Len(Trim$(de_informa.rsSel_NFsPancary.Fields("placaveic"))), " ") & _
                     Trim$(de_informa.rsSel_NFsPancary.Fields("placaveic"))
        Else
            xplacaveic = Mid$(Trim$(de_informa.rsSel_NFsPancary.Fields("placaveic")), 1, 3) & Mid$(Trim$(de_informa.rsSel_NFsPancary.Fields("placaveic")), 5, 4)
        End If

        xlinha = "1" & de_informa.rsSel_NFsPancary.Fields("remet_cgc") & xdata & "06460" & "00000" & Mid$(de_informa.rsSel_NFsPancary.Fields("filialctc"), 5, 6) & _
                 xplacaveic & "025" & zeros(de_informa.rsSel_NFsPancary.Fields("valmerc") * 100, 17) & "0000000000000000" & _
                 "0" & "99" & "0001" & String(11 - Len(SoNumeros(de_informa.rsSel_MotoristaPamcary.Fields("cpf"))), "0") & _
                 SoNumeros(de_informa.rsSel_MotoristaPamcary.Fields("cpf")) & Trim$(de_informa.rsSel_MotoristaPamcary.Fields("nome")) & _
                 String(50 - Len(Trim$(de_informa.rsSel_MotoristaPamcary.Fields("nome"))), " ") & "2" & "02"

        xcontador = xcontador + 1

        Print #1, xlinha
        
        If de_informa.rsSel_NFsPancary.Fields("uf_dest") = "AC" Then
            xcoddest = "001"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "AL" Then
            xcoddest = "002"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "AP" Then
            xcoddest = "003"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "AM" Then
            xcoddest = "004"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "BA" Then
            xcoddest = "005"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "CE" Then
            xcoddest = "006"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "DF" Then
            xcoddest = "007"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "ES" Then
            xcoddest = "008"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "GO" Then
            xcoddest = "009"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "MA" Then
            xcoddest = "010"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "MT" Then
            xcoddest = "011"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "MS" Then
            xcoddest = "012"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "MG" Then
            xcoddest = "013"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "PA" Then
            xcoddest = "014"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "PB" Then
            xcoddest = "015"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "PR" Then
            xcoddest = "016"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "PE" Then
            xcoddest = "017"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "PI" Then
            xcoddest = "018"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "RJ" Then
            xcoddest = "019"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "RN" Then
            xcoddest = "020"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "RS" Then
            xcoddest = "021"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "RO" Then
            xcoddest = "022"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "RR" Then
            xcoddest = "023"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "SC" Then
            xcoddest = "024"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "SP" Then
            xcoddest = "025"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "SE" Then
            xcoddest = "026"
        ElseIf de_informa.rsSel_NFsPancary.Fields("uf_dest") = "TO" Then
            xcoddest = "027"
        End If
        
        xlinha = "8" & xcoddest & zeros(de_informa.rsSel_NFsPancary.Fields("valmerc") * 100, 17)
        
        xcontador = xcontador + 1
        
        Print #1, xlinha
        
        de_informa.rsSel_NFsPancary.MoveNext
        
        DoEvents
    Loop
        
    xlinha = "999999999999999" & zeros(xcontador, 9)
    
    Print #1, xlinha
        
    Close #1
    
    MsgBox "Processo Finalizado !"

    
        
        

End Sub
