VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Samrt Easy Email"
   ClientHeight    =   2940
   ClientLeft      =   4260
   ClientTop       =   3645
   ClientWidth     =   5670
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About EasyHtml"
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2200
      Left            =   120
      ScaleHeight     =   2205
      ScaleWidth      =   5460
      TabIndex        =   1
      Top             =   120
      Width           =   5460
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         FillColor       =   &H0000C0C0&
         Height          =   1455
         Left            =   80
         ScaleHeight     =   1455
         ScaleWidth      =   5295
         TabIndex        =   3
         Top             =   680
         Width           =   5295
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For product information and help please email "
            Height          =   195
            Left            =   825
            TabIndex        =   5
            Top             =   435
            Width           =   3240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mjai_kumar@hotmail.com"
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   1575
            TabIndex        =   4
            Top             =   840
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   80
         ScaleHeight     =   585
         ScaleWidth      =   5265
         TabIndex        =   2
         Top             =   80
         Width           =   5295
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Smart Easy Mail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   1800
            TabIndex        =   6
            Top             =   180
            Width           =   1740
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ########################################################################
' #                                                                      #
' #             This is the About Box for Smart Easy Email                        #
' #                                                                      #
' #                                                                      #
' #                  Copyright 1999 Eric Banker                          #
' #                  All Rights Reserved                                 #
' #                                                                      #
' ########################################################################

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub


