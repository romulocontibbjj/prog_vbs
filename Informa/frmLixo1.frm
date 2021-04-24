VERSION 5.00
Begin VB.Form frmLixo1 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   2745
   ClientTop       =   1860
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6420
   Begin VB.CommandButton cmdProcessa 
      Caption         =   "Processar"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCadastra 
      Caption         =   "Cadastra"
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtCTC 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filial"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton opt03 
         Caption         =   "03"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton opt01 
         Caption         =   "01"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLixo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastra_Click()
    If de_informa.rslixo_selctc.State = 1 Then de_informa.rslixo_selctc.Close
    If opt01 = True Then
        xfilial = "01"
    Else
        xfilial = "03"
    End If
    de_informa.lixo_selctc xfilial & "00" & txtCTC
    If de_informa.rslixo_selctc.RecordCount > 0 Then
        MsgBox "FILIAL CTC JÁ CADASTRADO !"
        txtCTC.SetFocus
    Else
        de_informa.lixo_insctc xfilial & "00" & txtCTC, xfilial, txtCTC, ""
        txtCTC.SetFocus
    End If
End Sub

Private Sub cmdProcessa_Click()
    If de_informa.rslixo_seltodos.State = 1 Then de_informa.rslixo_seltodos.Close
    de_informa.lixo_seltodos
    Do Until de_informa.rslixo_seltodos.EOF
        xfilialctc = de_informa.rslixo_seltodos.Fields("filialctc")
        If de_informa.rslixo_SelOcorr.State = 1 Then de_informa.rslixo_SelOcorr.Close
        de_informa.lixo_SelOcorr xfilialctc
        If de_informa.rslixo_SelOcorr.RecordCount > 0 Then
            If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
            de_informa.Sel_Ctc_SAC xfilialctc
            xnfs = Trim$(de_informa.rsSel_Ctc_SAC.Fields("nfs"))
            de_informa.lixo_altctc xnfs, xfilialctc
        End If
        de_informa.rslixo_seltodos.MoveNext
    Loop
    MsgBox "Processo Finalizado ! "
End Sub

Private Sub txtCTC_GotFocus()
    txtCTC.SelStart = 0
    txtCTC.SelLength = 6
End Sub

Private Sub txtCTC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
