VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00400000&
   Caption         =   "COTY"
   ClientHeight    =   8235
   ClientLeft      =   3420
   ClientTop       =   2055
   ClientWidth     =   8805
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "*** *** *** COTY - Serviços de Entregas *** *** ***"
            TextSave        =   "*** *** *** COTY - Serviços de Entregas *** *** ***"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "15:34"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "31/12/2004"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1680
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1304
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3992
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":466C
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
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   1535
      ButtonWidth     =   2249
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CLIENTES"
            Description     =   "BOT_CLIENTES"
            Object.ToolTipText     =   "Cadastro de Clientes"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MOTOQUEIROS"
            Description     =   "Bot_Motoqueiros"
            Object.ToolTipText     =   "Cadastro de Motoqueiros"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MOTOCICLETAS"
            Description     =   "MOTOCICLETAS"
            Object.ToolTipText     =   "Cadastro de Motocicletas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Description     =   "EXIT"
            Object.ToolTipText     =   "Sair do Sistema"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu men_Aqruivoi 
      Caption         =   "&Arquivo"
      Begin VB.Menu men_Cadastro 
         Caption         =   "Efetuar &Cadastro"
         Begin VB.Menu men_Clientes 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu men_Motoqueiros 
            Caption         =   "&Motoqueiros"
         End
         Begin VB.Menu men_motocicletas 
            Caption         =   "M&otocicletas"
         End
      End
      Begin VB.Menu men_Traco 
         Caption         =   "-"
      End
      Begin VB.Menu men_Sair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu men_Pagamentos 
      Caption         =   "&Pagamentos"
   End
   Begin VB.Menu men_relatorios 
      Caption         =   "&Relatorios"
   End
   Begin VB.Menu men_Sair2 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub men_Clientes_Click()
'FRM_CadastrodeClientes.Show

End Sub

Private Sub men_motocicletas_Click()
FRM_CadastrodeMotos.Show



End Sub

Private Sub men_Motoqueiros_Click()
FRM_CadastrodeMotoqueiros.Show

End Sub

Private Sub men_Sair_Click()
Unload Me

End Sub

Private Sub men_Sair2_Click()
Unload Me

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1:
    men_Clientes_Click

Case 2:
    men_Motoqueiros_Click

Case 3:
    men_motocicletas_Click

Case 4:
    Unload Me

End Select






End Sub
