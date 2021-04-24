VERSION 5.00
Begin VB.Form frmEmissaoCANCAWB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar AWB"
   ClientHeight    =   4875
   ClientLeft      =   3600
   ClientTop       =   1620
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoCANCAWB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   315
      Left            =   3300
      TabIndex        =   4
      Top             =   300
      Width           =   855
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar AWB"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4380
      Width           =   1875
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4380
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Caption         =   "Motivo"
      Height          =   2475
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4035
      Begin VB.TextBox TxtMotivo 
         BackColor       =   &H00C0FFFF&
         Height          =   1845
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label LblCaract 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total de Caracteres Disp.: 250"
         Height          =   195
         Left            =   1755
         TabIndex        =   12
         Top             =   2160
         Width           =   2160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   780
      Width           =   4035
      Begin VB.Label LblStatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1965
         TabIndex        =   10
         Top             =   420
         Width           =   105
      End
   End
   Begin VB.Frame FraDadosManual 
      Caption         =   "Filial / Cia. / AWB / Dig."
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3075
      Begin VB.TextBox TxtSigla 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox TxtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox TxtAWB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtDig 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   3
         Top             =   240
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmEmissaoCANCAWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodAwb As String

Private Sub CmdBuscar_Click()

If Len(TxtFilial.Text) = 0 Or Len(TxtSigla.Text) = 0 Or Len(TxtAWB.Text) = 0 Or Len(TxtDig.Text) = 0 Then
Exit Sub
End If

Me.MousePointer = 11
DoEvents
CodAwb = TxtFilial.Text & TxtSigla.Text & String(10 - Len(Trim(Str(Val(TxtAWB.Text)))), "0") & Trim(Str(Val(TxtAWB.Text))) & Trim(Str(Val(TxtDig.Text)))

    If de_informa.rsConsultaAWB.State = 1 Then de_informa.rsConsultaAWB.Close
    de_informa.ConsultaAWB CodAwb
    
    If de_informa.rsConsultaAWB.RecordCount = 0 Then
    MsgBox "AWB não encontrado!", vbExclamation, ""
    Me.MousePointer = 0
    TxtAWB.SetFocus
    DoEvents
    Exit Sub
    End If

If de_informa.rsConsultaAWB.Fields("cancelado") <> "X" Or IsNull(de_informa.rsConsultaAWB.Fields("cancelado")) = True Then
LblStatus.Caption = "Disponível para Cancelamento"
CmdCancelar.Enabled = True
Else
LblStatus.Caption = "AWB já Cancelado"
CmdCancelar.Enabled = False
TxtMotivo.Text = de_informa.rsConsultaAWB.Fields("canc_motivo")
TxtAWB.SetFocus
End If
Me.MousePointer = 0
DoEvents

End Sub

Private Sub CmdCancelar_Click()
    If Len(Trim(TxtMotivo.Text)) = 0 Then
    MsgBox "É necessário informar o motivo do cancelamento para prosseguir.", vbExclamation, ""
    Exit Sub
    End If

Me.MousePointer = 11
DoEvents

de_informa.cn_informa.BeginTrans
de_informa.CancelaAWB xUsuario, UCase(Trim(TxtMotivo.Text)), DataHora("DATA"), DataHora("HORA"), CodAwb
de_informa.Alt_FormularioStatuscanc "C", xUsuario, UCase(Trim(TxtMotivo.Text)), TxtSigla.Text, TxtAWB.Text, TxtFilial.Text
de_informa.cn_informa.CommitTrans

TxtMotivo.Text = ""
'TxtFilial.Text = ""
'TxtSigla.Text = ""
TxtAWB.Text = ""
TxtDig.Text = ""
LblStatus.Caption = ""
CmdCancelar.Enabled = False

Me.MousePointer = 0
DoEvents
TxtAWB.SetFocus
End Sub

Private Sub CmdSair_Click()
Unload Me
End Sub

Private Sub TxtAWB_GotFocus()
TxtAWB.SelStart = 0
TxtAWB.SelLength = 200
End Sub

Private Sub TxtAWB_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtDig_GotFocus()
TxtDig.SelStart = 0
TxtDig.SelLength = 200
End Sub

Private Sub TxtDig_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtFilial_GotFocus()
TxtFilial.SelStart = 0
TxtFilial.SelLength = 200
End Sub

Private Sub TxtFilial_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtMotivo_Change()
Dim X As Integer
X = TxtMotivo.SelStart
TxtMotivo = UCase(TxtMotivo.Text)
TxtMotivo.SelStart = X
LblCaract.Caption = "Total de Caracteres Disp.: " & 250 - Len(TxtMotivo.Text)
DoEvents
End Sub

Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub TxtSigla_Change()
Dim X As Integer
X = TxtSigla.SelStart
TxtSigla.Text = UCase(TxtSigla.Text)
TxtSigla.SelStart = X
End Sub

Private Sub TxtSigla_GotFocus()
TxtSigla.SelStart = 0
TxtSigla.SelLength = 200
End Sub

Private Sub TxtSigla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub
