VERSION 5.00
Begin VB.Form MDI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trabalho de Bissecção - 2ºCCN"
   ClientHeight    =   2340
   ClientLeft      =   4230
   ClientTop       =   4245
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   5115
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   315
         Left            =   420
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmd_limpar 
         Caption         =   "&L IMPAR"
         Height          =   1755
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   195
      End
      Begin VB.CommandButton CMD_CALC 
         Caption         =   "&CALCULAR"
         Height          =   315
         Left            =   420
         TabIndex        =   9
         Top             =   1500
         Width           =   1455
      End
      Begin VB.TextBox TXT_B 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   420
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TXT_E 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TXT_A 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   420
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "K:"
         Height          =   315
         Left            =   3060
         TabIndex        =   17
         Top             =   1740
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "XM:"
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "A:"
         Height          =   255
         Left            =   3060
         TabIndex        =   15
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "B:"
         Height          =   315
         Left            =   3060
         TabIndex        =   14
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "E:"
         Height          =   375
         Left            =   3060
         TabIndex        =   13
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label LAB_E 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3360
         TabIndex        =   0
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label LAB_K 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3360
         TabIndex        =   1
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label LAB_B 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label LAB_M 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label LAB_A 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "E:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "B:"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_CALC_Click()
Dim XA As Double
Dim XB As Double
Dim XE As Double
Dim XM As Double
Dim I As Integer
Dim FA As Double
Dim FB As Double
Dim FM As Double
Dim XE_1 As Double
Dim K As Double
Dim QTD_E As Integer
Dim MDI As Double

XA = TXT_A.Text
XB = TXT_B.Text
XE = TXT_E.Text
XM = (XB + XA) / 2
XE_1 = XB - XA
QTD_E = Len(TXT_E.Text)
MDI = Mid(XE_1, 1, QTD_E)



Do While (MDI > XE)
FA = (1.8 * Exp(XA)) - (2 * Exp(-(XA) / 2))
FB = (1.8 * Exp(XB)) - (2 * Exp(-(XB) / 2))
FM = (1.8 * Exp(XM)) - (2 * Exp(-(XM) / 2))

If (FA * FM) < 0 Then
    XB = XM
ElseIf (FB * FM) < 0 Then
    XA = XM
End If

XM = (XB + XA) / 2
XE_1 = XB - XA
K = K + 1
MDI = Mid(XE_1, 1, QTD_E)


Loop


LAB_A.Caption = XA
LAB_B.Caption = XB
LAB_M.Caption = XM
LAB_E.Caption = XE_1
LAB_K.Caption = K

MsgBox "O Valor da Raiz de A: " & XA & Chr$(13) & "O Valor da Raiz de B:" & XB & Chr$(13) & "O Valor da Raiz da XM: " & XM & Chr$(13) & "Quantidade de Interações: " & K & Chr$(13) & "ERRO: " & XE_1, vbInformation, "FUNÇÃO: (1,8 * EXP(X)) - (2 * EXP(-(X) / 2)"





End Sub

Private Sub cmd_limpar_Click()
TXT_A.Text = Empty
TXT_B.Text = Empty
TXT_E.Text = Empty
LAB_A.Caption = Empty
LAB_B.Caption = Empty
LAB_E.Caption = Empty
LAB_M.Caption = Empty
LAB_K.Caption = Empty
TXT_A.SetFocus

End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub
