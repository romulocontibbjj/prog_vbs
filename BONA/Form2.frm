VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   2130
   ClientTop       =   1920
   ClientWidth     =   10785
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   10785
   Begin VB.TextBox txt_url 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmd_ir 
      Caption         =   "IR"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10575
      ExtentX         =   18653
      ExtentY         =   12938
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_ir_Click()
WebBrowser1.Navigate Trim$(txt_url)

End Sub
