VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmLogIn2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LogIn"
   ClientHeight    =   1110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1164
      BandCount       =   2
      Enabled         =   0   'False
      _CBWidth        =   3495
      _CBHeight       =   660
      _Version        =   "6.0.8169"
      Caption1        =   "Usuário:"
      Child1          =   "txtName"
      MinHeight1      =   285
      Width1          =   1965
      NewRow1         =   0   'False
      Caption2        =   "Senha:"
      Child2          =   "txtPassword"
      MinHeight2      =   285
      Width2          =   825
      NewRow2         =   -1  'True
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   810
         TabIndex        =   3
         Top             =   30
         Width           =   2595
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   735
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   345
         Width           =   2670
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1140
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogIn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OKButton_Click()
    Me.Hide
End Sub
