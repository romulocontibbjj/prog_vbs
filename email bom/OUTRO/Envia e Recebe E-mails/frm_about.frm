VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frank Mailer"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_about 
      Alignment       =   2  'Center
      Caption         =   "Thankyou for download Frankmailer, if you would like to know more or report a bug, please send an email to np24@blueyonder.co.uk"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   2400
      Width           =   3795
   End
   Begin VB.Image img_splash 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Left            =   60
      Picture         =   "frm_about.frx":0000
      Top             =   60
      Width           =   3780
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'FORM EVENTS
'**********************************
    Private Sub Form_Unload(Cancel As Integer)
        frm_main.Show
    End Sub

'**********************************
'PICTURE MOUSE CLICK
'**********************************
    Private Sub img_splash_Click()
        Unload Me
    End Sub
