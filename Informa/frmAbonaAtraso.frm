VERSION 5.00
Begin VB.Form frmAbonaAtraso 
   Caption         =   "Justificativa do Atraso"
   ClientHeight    =   2910
   ClientLeft      =   2655
   ClientTop       =   1500
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   6615
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtDiasJustif 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   1
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdProcessaJustif 
         Caption         =   "Justificar Atraso"
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtMotivoJustif 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   2880
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton cmdMenosUm 
         Caption         =   "-"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton cmdMaisUm 
         Caption         =   "+"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   840
         Width           =   525
      End
      Begin VB.Label lblPrazoContr 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Contratual:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   440
         Width           =   1215
      End
      Begin VB.Label lblPrazoReal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Real:"
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   440
         Width           =   825
      End
      Begin VB.Label lblDiasAtraso 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dias de Atraso:"
         Height          =   195
         Left            =   4320
         TabIndex        =   8
         Top             =   440
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dias para Justificar:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1150
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmAbonaAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdMaisUm_Click()
    If Val(txtDiasJustif) < Val(lblDiasAtraso) Then txtDiasJustif = Val(txtDiasJustif + 1)
End Sub

Private Sub cmdMenosUm_Click()
    If Val(txtDiasJustif) > 1 Then txtDiasJustif = Val(txtDiasJustif - 1)
End Sub

Private Sub cmdProcessaJustif_Click()
    If MsgBox("Confirma o Abono deste Atraso de Entrega ?", vbYesNo, "Confirmação") = vbYes Then
        de_informa.Alt_AbonoAtraso txtDiasJustif, xusuario, datahora("DATAHORA"), txtMotivoJustif, frmAnEntregas.gridAtrasos.Columns(0)
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "PROCESSO", xusuario, "Abono de atraso CTC: " & frmAnEntregas.gridAtrasos.Columns(0)
        MsgBox "Atraso Abonado / Justificado !"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbonaAtraso = Nothing
End Sub

Private Sub txtDiasJustif_GotFocus()
    txtDiasJustif.SelStart = 0
    txtDiasJustif.SelLength = 2
End Sub

Private Sub txtDiasJustif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtMotivoJustif_GotFocus()
    txtMotivoJustif.SelStart = 0
    txtMotivoJustif.SelLength = 40
End Sub

Private Sub txtMotivoJustif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtMotivoJustif_LostFocus()
    txtMotivoJustif = UCase(txtMotivoJustif)
End Sub
