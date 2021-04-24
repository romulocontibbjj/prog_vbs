VERSION 5.00
Begin VB.Form POWER 
   Caption         =   "Form4"
   ClientHeight    =   4215
   ClientLeft      =   3450
   ClientTop       =   2820
   ClientWidth     =   7395
   LinkTopic       =   "Form4"
   ScaleHeight     =   4215
   ScaleWidth      =   7395
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "POWER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim Word As New Word.Application
Dim wordDoc As Word.Document

Set Word = CreateObject("word.Application")
Set wordDoc = Word.Documents.Add(, newtemplate:=True)
Word.Visible = True






End Sub
