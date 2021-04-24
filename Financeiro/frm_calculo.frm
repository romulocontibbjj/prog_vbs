VERSION 5.00
Begin VB.Form frm_calculo 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Matematica Financeira - Basic"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_calculo.frx":0000
   ScaleHeight     =   5355
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Ccmd_jurosreal 
      Caption         =   "J&uros em Real"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmd_limpar 
      Caption         =   "&Zerar"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "SAI&R"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txt_fv 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txt_n 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txt_i 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "0"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txt_j 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "0"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt_pv 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_j 
      Caption         =   "&j"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmd_PV 
      Caption         =   "&PV"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmd_i 
      Caption         =   "&i"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmd_n 
      Caption         =   "&n"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmd_FV 
      Caption         =   "&FV"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "FV:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "N:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "I:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "J:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "PV:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "frm_calculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ccmd_jurosreal_Click()
If txt_i <= 1 Then
    Dim x As Currency
    x = txt_i.Text * 100
    txt_i.Text = x
   
Else
    txt_i.Text = txt_i / 100
End If

End Sub

Private Sub cmd_FV_Click()
Dim pv As Currency
Dim i_n2 As Currency
pv = Val(txt_pv.Text)
If txt_i = 0 Then
    txt_fv.Text = pv + Val(txt_j)
Else
    i_n2 = 1 + (txt_i * txt_n)
    txt_fv.Text = pv * i_n2
End If


End Sub

Private Sub cmd_i_Click()
Dim j As Single
Dim pv_n As Single
Dim fv_pv As Single
Dim n As Single


If txt_j = 0 Then
    fv_pv = (txt_fv / txt_pv) - 1
    n = fv_pv / txt_n
    txt_i = Format(n, "##0.0000")
    
Else
j = Val(txt_j.Text)
pv_n = txt_pv * txt_n
    txt_i.Text = j / pv_n
End If
    


End Sub

Private Sub cmd_j_Click()
Dim j As Currency

If txt_i = 0 Then
    txt_j.Text = txt_fv - txt_pv
Else
    j = txt_pv * txt_i * txt_n
    txt_j.Text = j
End If


End Sub

Private Sub cmd_limpar_Click()
txt_fv = 0
txt_i = 0
txt_j = 0
txt_n = 0
txt_pv = 0
txt_pv.SetFocus
End Sub

Private Sub cmd_n_Click()
Dim j As Currency
Dim pv_i As Currency
j = Val(txt_j.Text)
pv_i = txt_pv * txt_i
If j Or pv_i = 0 Then
    MsgBox "Falta-se Dados, confira os campos PV, J e I", vbExclamation, "Confira Dados"
    txt_pv.SetFocus
Else
    txt_n.Text = j / pv_i
End If



End Sub

Private Sub cmd_PV_Click()
Dim j As Currency
Dim i_n As Currency

If txt_i = 0 Then
    txt_pv.Text = txt_fv - txt_j
Else
    j = Val(txt_j.Text)
    i_n = txt_i * txt_n
    txt_pv.Text = j / i_n
End If
txt_pv.Text = Format(txt_pv, "000.000")



End Sub

Private Sub cmd_sair_Click()
If MsgBox("Vc tem certeza que Deseja Sair???", vbQuestion + vbOKCancel, "SAIR???") = vbOK Then
    Unload frm_calculo
    frm_Bye.Show
  End If
    
    

End Sub

Private Sub txt_pv_DblClick()
MsgBox "Valor Inicial" & txt_pv, vbInformation, "Valor Inicial"


End Sub
