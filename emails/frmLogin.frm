VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LogIn"
   ClientHeight    =   1470
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2775
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   7
      Top             =   1095
      Width           =   2775
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções:"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox Check3 
         Caption         =   "Não Exibir Mais Essa Tela"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Salvar Informações"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Exigir Autenticação"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   435
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1164
      BandCount       =   2
      Enabled         =   0   'False
      _CBWidth        =   2775
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
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   735
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   345
         Width           =   1950
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   810
         TabIndex        =   1
         Top             =   30
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    Check2.Enabled = True
    CoolBar1.Enabled = True
    txtName.Enabled = True
    txtPassword.Enabled = True
    Me.Height = 2505
Else
    Check2.Enabled = False
    CoolBar1.Enabled = False
    txtName.Enabled = False
    txtPassword.Enabled = False
    Me.Height = 1845
End If
End Sub

Private Sub OKButton_Click()
If Check1.Value = vbChecked Then
    SaveSetting "VBSender", "Geral", "Exigir Autenticação", True
Else
    SaveSetting "VBSender", "Geral", "Exigir Autenticação", False
End If
If Check2.Value = vbChecked Then
    SaveSetting "VBSender", "Geral", "Salvo", True
    SaveSetting "VBSender", "Geral", "User", txtName.Text
    SaveSetting "VBSender", "Geral", "Pass", txtPassword.Text
Else
    SaveSetting "VBSender", "Geral", "Salvo", False
End If
If Check3.Value = vbChecked Then
    SaveSetting "VBSender", "Geral", "Exibir", False
Else
    SaveSetting "VBSender", "Geral", "Exibir", True
End If
Unload Me
End Sub
