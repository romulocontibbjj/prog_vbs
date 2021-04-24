VERSION 5.00
Begin VB.Form frm_cadastro 
   BackColor       =   &H00000040&
   Caption         =   "CADASTRO"
   ClientHeight    =   3540
   ClientLeft      =   2340
   ClientTop       =   2295
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6525
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "Proximo"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmd_abr1 
      Caption         =   "ABR"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd_abrir 
      Caption         =   "Bloco de Notas"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton CMD_GRAVAR 
      Caption         =   "GRAVAR"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txt_end 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txt_nome 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "ENDEREÇO:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   "NOME:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frm_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nome() As String
Private endereco() As String

Private Sub cmd_abrir_Click()
Shell "notepad.exe c:\cadastro.txt", vbNormalFocus

End Sub

Private Sub CMD_GRAVAR_Click()

If Trim(txt_nome.Text) = Empty Or Trim(txt_end.Text) = Empty Then
    If txt_nome.Text = Empty Then
        MsgBox "Nome Inválido", vbInformation, "NOME"
        txt_nome.SetFocus
    Else
        MsgBox "Endereço inválido", vbInformation, "ENREÇO"
        txt_end.SetFocus
        End If
Else

Open "C:\CADASTRO.TXT" For Append As #1
Print #1, UCase(txt_nome.Text) & "#" & UCase(txt_end.Text)
Close #1

MsgBox "Cliente " & UCase(txt_nome.Text) & " Cadastrado", vbInformation, "Cadastro OK!"

txt_end.Text = Empty
txt_nome.Text = Empty
txt_nome.SetFocus

End If


End Sub

Private Sub cmd_abr1_Click()
Dim vcliente As String
Dim x As Integer
Dim i As Integer
Open "C:\CADASTRO.TXT" For Input As #1
Do While EOF(1) = False
Line Input #1, vcliente
x = InStr(vcliente, "#")
ReDim Preserve nome(i), endereco(i)
nome(i) = Mid(vcliente, 1, x - 1)
endereco(i) = Mid(vcliente, x + 1)
i = i + 1
Loop
Close #1
MsgBox vcliente




End Sub


Private Sub cmd_proximo_Click()
Static vpos As Integer

txt_nome.Text = nome(vpos)
txt_end.Text = endereco(vpos)
vpos = vpos + 1

If vpos > UBound(nome) Then
    vpos = 0
End If


End Sub
