VERSION 5.00
Begin VB.Form frm_listagem 
   Caption         =   "Listagem de Produtos e Promoções"
   ClientHeight    =   4905
   ClientLeft      =   3945
   ClientTop       =   3450
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   6465
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_produto 
      Caption         =   "Produtos"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd_promocao 
      Caption         =   "Promoção"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Saindu!!!!!!!!!!!!!!!!!"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmd_1 
      Caption         =   ">"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmd_2 
      Caption         =   ">>"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmd_3 
      Caption         =   "<"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmd_4 
      Caption         =   "<<"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.ListBox lst_Promocoes 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   4320
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ListBox lst_produtos 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txt_produto 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Promoções"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Produtos"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frm_listagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_1_Click()
If lst_produtos.ListCount = 0 Then
    MsgBox "Ae Comédia...Não sei se vc percebeu mas a lista esta vazia!!!!", vbCritical, "BOBÃO"
ElseIf lst_produtos.ListIndex = -1 Then
    MsgBox "Selecione um produto", vbInformation, "Produto"
Else
lst_Promocoes.AddItem lst_produtos.List(lst_produtos.ListIndex)
lst_produtos.RemoveItem lst_produtos.ListIndex
End If

End Sub

Private Sub cmd_2_Click()
Dim i As Integer
Dim x As Integer

If lst_produtos.ListCount = 0 Then
MsgBox "Cara.....sse liga não tem produto!!!!!!!!!!!!!!!!!", vbCritical, "MANÉZÃO!!!!"
txt_produto.SetFocus
Else
x = lst_produtos.ListCount - 1

For i = 0 To x Step 1
lst_Promocoes.AddItem lst_produtos.List(i)
Next
lst_produtos.Clear
End If


End Sub

Private Sub cmd_3_Click()


If lst_Promocoes.ListCount = 0 Then
    MsgBox "Não há dados a Serem movidos", vbExclamation, "ERRO"
ElseIf lst_Promocoes.ListIndex = -1 Then
    MsgBox "Selecione um produto", vbInformation, "Produto"
Else
lst_produtos.AddItem lst_Promocoes.List(lst_Promocoes.ListIndex)
lst_Promocoes.RemoveItem lst_Promocoes.ListIndex
End If

End Sub

Private Sub cmd_4_Click()
Dim i As Integer
Dim x As Integer

If lst_Promocoes.ListCount = 0 Then
MsgBox "Cara.....sse liga não tem produto!!!!!!!!!!!!!!!!!", vbCritical, "MANÉZÃO!!!!"
txt_produto.SetFocus
Else
x = lst_Promocoes.ListCount - 1

For i = 0 To x Step 1
lst_produtos.AddItem lst_Promocoes.List(i)
Next
lst_Promocoes.Clear
End If


End Sub

Private Sub cmd_ok_Click()
Dim x As Integer
Dim i As Integer
x = lst_produtos.ListCount
If Trim(txt_produto.Text) = Empty Then
    MsgBox "Digite um Produto", vbInformation, "Produto"
Else

x = lst_produtos.ListCount
For i = 0 To x Step 1
If Trim(txt_produto.Text) = lst_produtos.List(i) Then
MsgBox "NAO"
txt_produto.SelStart = 0
txt_produto.SelLength = Len(Trim(txt_produto.Text))
txt_produto.SetFocus
Exit Sub

End If

Next

 
     lst_produtos.AddItem txt_produto.Text
    txt_produto.Text = Empty
    txt_produto.SetFocus
End If



End Sub

Private Sub cmd_produto_Click()
Dim vtexto As String
Open "C:\PRODUTO.TXT" For Input As #1
Do While EOF(1) = False



End Sub

Private Sub cmd_promocao_Click()
Dim vtexto As String
Dim i As Integer
Dim x As Integer
'x = lst_Promocoes.ListCount - 1
Open "C:\PROMOCAO.txt" For Output As #1
Do While x < lst_Promocoes.ListCount
Print #1, lst_Promocoes.List(x)
x = x + 1
Loop
Close #1

    
    


End Sub

Private Sub cmd_sair_Click()
Unload Me
End Sub
