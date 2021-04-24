VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   2715
   ClientTop       =   2175
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6375
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   975
      Left            =   3360
      TabIndex        =   7
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Verifica se Existe Arq ou Pasta"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "COPIA E REMOVE"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REMOVE PASTA"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CRIA PASTA"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'CRIA NOVA PASTA

MkDir (Dir1.Path & "\" & InputBox("Nome do Arquivo", TESTE))
Me.Refresh

End Sub

Private Sub Command2_Click()
'REMOVE PASTA

RmDir (Dir1.Path)
Me.Refresh

'KILL(INSTRUÇÃO) - APAGA ARQUIVO


End Sub

Private Sub Command3_Click()
FileCopy Dir1.Path & "\" & File1.FileName, "D:\AVISA\" & File1.FileName
Kill Dir1.Path & "\" & File1.FileName



End Sub

Private Sub Command4_Click()
    Kill "H:\Tecsidel\Colinas\NewReclamaPeriodo\tmpftp\*.*"
    
End Sub

Private Sub Command5_Click()
If Xbusca(UCase(Text1.Text), True) = True Then
    MsgBox "OK"
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
Dir1.Path = "D:\"
Dir1.Refresh
File1.Path = Dir1.Path
End Sub
