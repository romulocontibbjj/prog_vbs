VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frm_progress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<task title>"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pro_progress 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   420
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image img_progress 
      Height          =   480
      Left            =   240
      Picture         =   "frm_progress.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbl_step 
      Caption         =   "<step>"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frm_progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'FORM EVENTS
'**********************************
    Private Sub Form_Unload(Cancel As Integer)
        If mailbusy Then
            Cancel = 1
            Me.Hide
        End If
    End Sub
