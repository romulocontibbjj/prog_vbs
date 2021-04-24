VERSION 5.00
Begin VB.Form frm_Misturador 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misturador"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frm_Misturador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Voltar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VOLTAR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      MouseIcon       =   "frm_Misturador.frx":0442
      Picture         =   "frm_Misturador.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.HScrollBar Hsb_Red 
      Height          =   255
      Left            =   360
      Max             =   255
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.HScrollBar Hsb_Green 
      Height          =   255
      Left            =   360
      Max             =   255
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.HScrollBar Hsb_Blue 
      Height          =   255
      Left            =   360
      Max             =   255
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Lab_red 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Lab_Green 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Lab_Blue 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lab_Total 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frm_Misturador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Voltar_Click()
Hsb_Blue.Value = 255
Hsb_Green.Value = 255
Hsb_Red.Value = 255
frm_Misturador.BackColor = &H404040





End Sub

Private Sub Hsb_Blue_Change()
Lab_Blue.BackColor = RGB(0, 0, Hsb_Blue)
Lab_Blue.Caption = Hsb_Blue.Value
lab_Total.BackColor = RGB(Hsb_Red, Hsb_Green, Hsb_Blue)

End Sub

Private Sub Hsb_Blue_Scroll()
Hsb_Blue_Change
End Sub

Private Sub Hsb_Green_Change()
Lab_Green.BackColor = RGB(0, Hsb_Green, 0)
Lab_Green.Caption = Hsb_Green.Value
lab_Total.BackColor = RGB(Hsb_Red, Hsb_Green, Hsb_Blue)

End Sub

Private Sub Hsb_Green_Scroll()
Hsb_Green_Change
End Sub

Private Sub Hsb_Red_Change()
Lab_red.BackColor = RGB(Hsb_Red, 0, 0)
Lab_red.Caption = Hsb_Red.Value
lab_Total.BackColor = RGB(Hsb_Red, Hsb_Green, Hsb_Blue)

End Sub

Private Sub Hsb_Red_Scroll()
Hsb_Red_Change
End Sub

Private Sub lab_Total_DblClick()
Me.BackColor = lab_Total.BackColor
End Sub
