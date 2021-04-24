VERSION 5.00
Begin VB.Form frm_pendencia 
   Caption         =   "Pendências da hoje"
   ClientHeight    =   8235
   ClientLeft      =   4470
   ClientTop       =   1845
   ClientWidth     =   7275
   Icon            =   "frm_pendencia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7275
   Begin VB.Frame frm_diario 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   7800
         Width           =   615
      End
      Begin VB.CommandButton cmd_gerar_arq 
         Caption         =   "Gerar Arquivo"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   7320
         Width           =   2295
      End
      Begin VB.CommandButton cmd_fechar 
         Caption         =   "FECHAR"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   7320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_descr 
         Caption         =   "OK"
         Height          =   255
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt_descr 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3840
         Left            =   4200
         TabIndex        =   6
         Top             =   3360
         Width           =   2775
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3840
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   3375
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1680
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label4 
         Caption         =   "DESCRIÇÃO:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "HORÁRIO"
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "FECHADOS"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ABERTO"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   3840
         X2              =   3840
         Y1              =   3360
         Y2              =   7200
      End
   End
End
Attribute VB_Name = "frm_pendencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_descr_Click()
List1.AddItem UCase(txt_descr.Text) & " - (" & Time & ")"
txt_descr.Text = Empty
txt_descr.SetFocus


End Sub

Private Sub cmd_fechar_Click()
If List1.ListIndex = -1 Then
    MsgBox "Selecione A Pendência a Ser Finalizada", vbInformation, "FECHAMENTO"
    
End If

List2.AddItem List1.List(List1.ListIndex)
List3.AddItem Time

List1.RemoveItem List1.ListIndex


'lst_Promocoes.AddItem lst_produtos.List(lst_produtos.ListIndex)
'lst_produtos.RemoveItem lst_produtos.ListIndex

End Sub

Private Sub cmd_gerar_arq_Click()
Dim xcontaberto As Integer
Dim xreg As String


xcontaberto = List1.ListCount

Open "C:\PENDECIAS_" & Day(Date) & Month(Date) & Hour(Time) & Minute(Time) & ".TXT" For Output As #1
xreg = "ABERTOS"
Print #1, xreg
Do Until xcontaberto = -1

xreg = List1.List(xcontaberto)

Print #1, xreg

xcontaberto = xcontaberto - 1

Loop

Close #1

MsgBox "ok"



End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Caption = "Pendências da hoje - " & Date

End Sub
