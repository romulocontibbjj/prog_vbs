VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1830
   ClientLeft      =   3360
   ClientTop       =   4455
   ClientWidth     =   6450
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1380
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label LblMensagem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6195
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CmdSair_Click()
Unload Me
End Sub
