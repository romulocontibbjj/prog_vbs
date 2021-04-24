VERSION 5.00
Begin VB.Form frmConexao 
   Caption         =   "Conexão ao Banco de Dados"
   ClientHeight    =   2730
   ClientLeft      =   1275
   ClientTop       =   1605
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   2730
   ScaleWidth      =   6015
   Begin VB.TextBox txtNomeBancoGrupo 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtNomeBanco 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtBanco 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdGravar 
      BackColor       =   &H80000009&
      Caption         =   "Gravar"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Nome do Banco de Dados Operação em GRUPO ( Se Houver ) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   3120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Banco de Dados Padrão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entre com o Servidor de Banco de Dados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "frmConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravar_Click()
    Open "c:\informa.cnx" For Output As #1
    Print #1, "CNX=" & txtBanco
    Print #1, "BCO=" & txtNomeBanco
    Print #1, "BGR=" & txtNomeBancoGrupo
    Close #1
    MsgBox "Conexão Gravada !", vbInformation
    Unload Me
End Sub

Private Sub Form_Load()
    
    If Dir("C:\informa.cnx") <> "" Then
    
        Open "C:\informa.cnx" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "CNX" Then
                txtBanco = Trim$(Mid$(xlinha, 5))
            End If
            If Mid$(xlinha, 1, 3) = "BCO" Then
                txtNomeBanco = Trim$(Mid$(xlinha, 5))
            End If
            If Mid$(xlinha, 1, 3) = "BGR" Then
                txtNomeBancoGrupo = Trim$(Mid$(xlinha, 5))
            End If
        Loop
        
        Close #1

    End If

End Sub
Private Sub txtBanco_GotFocus()
    txtBanco.SelStart = 0
    txtBanco.SelLength = Len(txtBanco)
End Sub
Private Sub txtNomeBanco_GotFocus()
    txtNomeBanco.SelStart = 0
    txtNomeBanco.SelLength = Len(txtNomeBanco)
End Sub
Private Sub txtNomeBanco_LostFocus()
    txtNomeBanco.Text = UCase(Trim$(txtNomeBanco))
End Sub
Private Sub txtNomeBancoGrupo_GotFocus()
    txtNomeBancoGrupo.SelStart = 0
    txtNomeBancoGrupo.SelLength = Len(txtNomeBancoGrupo)
End Sub
Private Sub txtNomeBancoGrupo_LostFocus()
    txtNomeBancoGrupo.Text = UCase(Trim$(txtNomeBancoGrupo))
End Sub
