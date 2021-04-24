VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   1920
   ClientTop       =   2340
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5175
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Image1.Picture = LoadPicture("C:\FOTOS\" & Frm_Diversos.lst_fotos.Text)
DoEvents

End Sub
