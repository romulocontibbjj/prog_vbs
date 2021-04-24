VERSION 5.00
Begin VB.Form frmObsOcorr 
   Caption         =   "Observação de Ocorrência"
   ClientHeight    =   3135
   ClientLeft      =   1905
   ClientTop       =   1875
   ClientWidth     =   7815
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   7815
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   2160
      Width           =   3690
      Begin VB.CommandButton cmdGravaN 
         Caption         =   "Sair"
         Height          =   495
         Left            =   1995
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdGravaS 
         Caption         =   "Gravar"
         Height          =   495
         Left            =   210
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtObs_Ocorr 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   120
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Label lblObsOcorr 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "de 300 (Max)"
      Height          =   195
      Left            =   2295
      TabIndex        =   6
      Top             =   2505
      Width           =   930
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   195
      Left            =   1875
      TabIndex        =   5
      Top             =   2505
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Caracteres: "
      Height          =   195
      Left            =   405
      TabIndex        =   4
      Top             =   2505
      Width           =   1485
   End
End
Attribute VB_Name = "frmObsOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGravaN_Click()
    Unload Me
End Sub
Private Sub cmdGravaS_Click()
    Dim xTxtObs_Ocorr As String, xc As Long
    For xc = 1 To Len(txtObs_Ocorr)
        If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM,./ 1234567890-\ÁÉÍÓÚÂÊÔÃÕÇÀÜÏË?><=+*()@%&;:", UCase(Mid$(txtObs_Ocorr, xc, 1)), vbTextCompare) = 0 Then
            xTxtObs_Ocorr = xTxtObs_Ocorr & " "
        Else
            xTxtObs_Ocorr = xTxtObs_Ocorr & Mid$(txtObs_Ocorr, xc, 1)
        End If
    Next
    txtObs_Ocorr = xTxtObs_Ocorr
    If frmPod.chkObsOcorr.Value = 1 Then
        de_informa.alt_obs_ocorr transctc(frmPod.txtfilial.Text, frmPod.txtCTC.Text), frmPod.GridOcorr.Columns(2), _
        frmPod.GridOcorr.Columns(0), frmPod.GridOcorr.Columns(1), lblObsOcorr.Caption & Trim$(txtObs_Ocorr.Text) & "."
        'atualiza o grid de ocorrências
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
        de_informa.Sel_ConsOcorr2 transctc(frmPod.txtfilial, frmPod.txtCTC), "01"
        Set frmPod.GridOcorr.DataSource = de_informa
        frmPod.GridOcorr.DataMember = "Sel_ConsOcorr2"
        frmPod.GridOcorr.Refresh
        lblObsOcorr.Caption = frmPod.GridOcorr.Columns(6)
        txtObs_Ocorr.Text = ""
    ElseIf frmPod.chkObsEntr.Value = 1 Then
        de_informa.alt_obs_entr transctc(frmPod.txtfilial.Text, frmPod.txtCTC.Text), "01", lblObsOcorr.Caption & Trim$(txtObs_Ocorr.Text) & "."
        If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
        de_informa.Sel_ConsOcorr transctc(frmPod.txtfilial, frmPod.txtCTC), "01"
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
            If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) Then lblObsOcorr.Caption = de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")
        End If
        txtObs_Ocorr.Text = ""
    End If
End Sub
Private Sub Form_Load()
    If frmPod.chkObsOcorr.Value = 1 Then
        lblObsOcorr.Caption = frmPod.GridOcorr.Columns(6)
    ElseIf frmPod.chkObsEntr.Value = 1 Then
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
            If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) Then lblObsOcorr.Caption = de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")
        End If
    End If
    lblChar.Caption = Len(lblObsOcorr.Caption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmObsOcorr = Nothing
End Sub

Private Sub txtObs_Ocorr_Change()
    lblChar.Caption = Len(lblObsOcorr.Caption) + Len(txtObs_Ocorr.Text)
End Sub

Private Sub txtObs_Ocorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
