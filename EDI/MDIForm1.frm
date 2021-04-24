VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00404040&
   Caption         =   "Geração de EDI´S"
   ClientHeight    =   8235
   ClientLeft      =   2550
   ClientTop       =   2505
   ClientWidth     =   9225
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5F4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   1535
      ButtonWidth     =   1323
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SISTEMA"
            Description     =   "Sistema"
            Object.ToolTipText     =   "Inica Sistema de Envio"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EDI´s"
            Description     =   "EDI"
            Object.ToolTipText     =   "Cadastro de EDI´s"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EMAILS"
            Description     =   "EMAILS"
            Object.ToolTipText     =   "Cadastro de Emails"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SAIR"
            Description     =   "SAIR"
            Object.ToolTipText     =   "Sair do Sistema"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mem_arq 
      Caption         =   "&Arquivo"
      Begin VB.Menu mem_Altera 
         Caption         =   "&Altera Edi"
      End
   End
   Begin VB.Menu mem_cad 
      Caption         =   "&Cadastros"
      Begin VB.Menu mem_cadEmail 
         Caption         =   "&Emails"
      End
      Begin VB.Menu mem_cadEnvios 
         Caption         =   "&Envios"
         Begin VB.Menu mem_cadastrados 
            Caption         =   "Cadastrados"
         End
      End
   End
   Begin VB.Menu mem_sistemas 
      Caption         =   "&Sistema"
   End
   Begin VB.Menu mem_sair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mem_Altera_Click()

frm_limpaOcorr.Show 1


End Sub

Private Sub mem_cadastrados_Click()
frm_Cadastrados.Show
End Sub

Private Sub mem_cadEmail_Click()
frm_Emails.Show 1

End Sub


Private Sub mem_sair_Click()
Unload Me

End Sub

Private Sub mem_sistemas_Click()

frm_verifica.Show

Unload MDIForm1


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1:
mem_sistemas_Click

Case 2:
mem_cadastrados_Click

Case 3:
mem_cadEmail_Click

Case 4:

mem_sair_Click


End Select



End Sub
