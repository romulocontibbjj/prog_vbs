VERSION 5.00
Begin VB.Form frm_Bye 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BYE BYE...BABY..."
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_Bye.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_bye 
      Caption         =   "Bye"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Bye, Bye...."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Bye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_bye_Click()
Unload frm_calculo
Unload Me

End Sub
