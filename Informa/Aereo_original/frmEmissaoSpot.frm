VERSION 5.00
Begin VB.Form frmEmissaoSpot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Defina os Parâmetros para Spot"
   ClientHeight    =   1215
   ClientLeft      =   3615
   ClientTop       =   3270
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoSpot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "Continuar"
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   780
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   1935
   End
   Begin VB.TextBox TxtAutorizador 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1260
      MaxLength       =   255
      TabIndex        =   1
      Top             =   360
      Width           =   2955
   End
   Begin VB.TextBox TxtKilo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Autorizador"
      Height          =   195
      Left            =   1260
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor por Kilo"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   930
   End
End
Attribute VB_Name = "frmEmissaoSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdContinuar_Click()
Dim xFrete, xPesoTx, xVlKilo As Currency
If Len(Trim(TxtKilo.Text)) = 0 Then
MsgBox "Você não informou o Valor do Kilo para a Tarifa Spot!", vbExclamation, ""
Exit Sub
ElseIf Len(Trim(TxtAutorizador.Text)) = 0 Then
MsgBox "Você não informou o nome do Autorizador para esta Terifa Spot!", vbExclamation, ""
Exit Sub
End If

    If (Val(SemPonto(frmEmissao.TxtPesoReal.Text)) / 10) > (Val(SemPonto(frmEmissao.TxtPesoCubado.Text)) / 10) Then
    xPesoTx = (Val(SemPonto(frmEmissao.TxtPesoReal.Text)) / 10)
    Else
    xPesoTx = (Val(SemPonto(frmEmissao.TxtPesoCubado.Text)) / 10)
    End If

xVlKilo = (Val(SemPonto(TxtKilo.Text)) / 10)
xFrete = xPesoTx * xVlKilo

frmEmissao.TxtFreteNacional.Text = xFrete * 100
    
    With frmEmissao
    .TxtFreteTotal.Text = (CDbl(.TxtFreteNacional.Text) + CDbl(.TxtFreteRegional.Text) + CDbl(.TxtTXOrigem.Text) + CDbl(.TxtTXDestino.Text) + CDbl(.TxtTXRedesp.Text) + CDbl(.TxtOutros1.Text) + CDbl(.TxtOutros2.Text)) * 100
    End With
Unload Me
End Sub

Private Sub TxtAutorizador_Change()
If Len(Trim(TxtAutorizador.Text)) > 0 Then
TxtAutorizador.Text = UCase(TxtAutorizador.Text)
TxtAutorizador.SelStart = Len(TxtAutorizador.Text)
End If
End Sub

Private Sub TxtAutorizador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtKilo_Change()
Call TextMoneyBox_Change(TxtKilo)
End Sub

Private Sub TxtKilo_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(TxtKilo)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtKilo_GotFocus()
Call TextMoneyBox_GotFocus(TxtKilo)
End Sub
