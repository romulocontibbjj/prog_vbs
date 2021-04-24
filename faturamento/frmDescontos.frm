VERSION 5.00
Begin VB.Form frmDescontos 
   Caption         =   "Conceder Desconto"
   ClientHeight    =   4050
   ClientLeft      =   1050
   ClientTop       =   1320
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7440
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdGravarDesconto 
      Caption         =   "Gravar Desconto"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Desconto / Abatimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   3015
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial-Fatura:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   840
         End
         Begin VB.Label lblFilialFatura 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtObsAbat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtAbat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblObsAcres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   20
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label lblAcrescimo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto da Fatura:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "( - )  Abatimento:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "( + )  Acréscimos:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "( = )  Valor da Fatura:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto com ICMS:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "( - ) Desc. ICMS:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblValorFaturaBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblValorFatura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblTipoAbat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   1395
         Width           =   3615
      End
      Begin VB.Label lblValorFaturaBrutoICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblValorICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   780
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmDescontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub

Private Sub cmdGravarDesconto_Click()

    If Len(Trim$(txtAbat)) = 0 Then txtAbat = "0"
    
    If (CDbl(SoNumeros(txtAbat)) / 100) <= 0 Then
        MsgBox "Valor de Desconto Inválido !"
        txtAbat.SetFocus
        Exit Sub
    End If
    
    If (CDbl(SoNumeros(txtAbat)) / 100) >= (CDbl(SoNumeros(lblValorFaturaBruto)) / 100) Then
        MsgBox "Valor de Desconto Inválido !"
        txtAbat.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Você Confirma o Desconto Nesta Fatura ?", vbYesNo + vbQuestion, "Desconto") = vbYes Then
        de_informa.Alt_DescontoFatura CDbl(SoNumeros(txtAbat)) / 100, lblTipoAbat, txtObsAbat, CDbl(SoNumeros(txtAbat)) / 100, xusuario, datahora("DATAHORA"), lblFilialFatura
        de_informa.Ins_FaturaHistorico lblFilialFatura, xusuario, "DESCONTO", txtObsAbat, ""
        Unload Me
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub txtAbat_Change()
    If Not IsNumeric(txtAbat) Then
        SendKeys "{BACKSPACE}"
        Exit Sub
    End If
    Call TextMoneyBox_Change(txtAbat)
    DoEvents
    If Len(Trim$(txtAbat)) > 0 Then
        If CDbl(SoNumeros(txtAbat)) / 100 > 0 Then
            txtObsAbat.Enabled = True
            txtObsAbat.BackColor = xamarelo1
            DoEvents
            Exit Sub
        End If
    End If
    txtObsAbat.Enabled = False
    txtObsAbat.BackColor = xbranco
    DoEvents

End Sub
Private Sub txtAbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAbat_LostFocus()
    If Len(Trim$(txtAbat)) > 0 Then
        If CDbl(SoNumeros(txtAbat)) / 100 > 0 Then
            lblValorFatura = Format((CDbl(SoNumeros(lblValorFaturaBruto)) / 100) - (CDbl(SoNumeros(txtAbat)) / 100), "##,###,##0.00")
            frmMotivosDesconto.Caption = "Conceder Desconto"
            frmMotivosDesconto.Show 1
            Exit Sub
        End If
    End If
End Sub

Private Sub txtObsAbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtObsAbat_LostFocus()
    txtObsAbat = UCase(Trim$(txtObsAbat))
End Sub

