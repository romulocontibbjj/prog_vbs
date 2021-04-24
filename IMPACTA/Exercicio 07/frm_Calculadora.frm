VERSION 5.00
Begin VB.Form frm_Calculadora 
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   2010
   ClientTop       =   2190
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   5940
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txt_valor2 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cbo_Oper 
      Height          =   315
      ItemData        =   "frm_Calculadora.frx":0000
      Left            =   1680
      List            =   "frm_Calculadora.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txt_valor1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lab_total 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "frm_Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vvalor1 As Integer
Private vvalor2 As Integer

Private Sub cmd_ok_Click()
Dim vtotal As Double
Dim i As Integer
Select Case cbo_Oper.Text
Case "+"
vtotal = vvalor1 + vvalor2
Case "-"
vtotal = vvalor1 - vvalor2
Case "*"
vtotal = vvalor1 * vvalor2
Case "/"
If vvalor1 = Empty Or vvalor2 = Empty Then
MsgBox "Não exite número dividido por 0 (ZERO)"
Else
vtotal = vvalor1 / vvalor2
End If
Case "!"
Dim x As Integer
vtotal = vvalor1
i = (vtotal - 1)
Do While i <> 0
vtotal = vtotal * i
i = i - 1
Loop

Case "^"
vtotal = vvalor1 ^ vvalor2
Case "MOD"
vtotal = vvalor1 Mod vvalor2
End Select

If vtotal < 0 Then
lab_total.ForeColor = vbRed
ElseIf vtotal > 0 Then
lab_total.ForeColor = vbBlue
ElseIf vtotal = 0 Then
lab_total.ForeColor = vbBlack
End If
lab_total.Caption = vtotal

End Sub

Private Sub txt_valor1_LostFocus()
vvalor1 = Val(txt_valor1.Text)

End Sub

Private Sub txt_valor2_LostFocus()
vvalor2 = Val(txt_valor2.Text)

End Sub
