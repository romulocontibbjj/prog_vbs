VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   "COTY"
   ClientHeight    =   8205
   ClientLeft      =   2985
   ClientTop       =   2145
   ClientWidth     =   10875
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7830
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "*** COTY MOTOS - Mensageiros motorizados s/c Ltda - ME ***"
            TextSave        =   "*** COTY MOTOS - Mensageiros motorizados s/c Ltda - ME ***"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "15:17"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "07/01/2005"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3592
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":426C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4F46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1111
      ButtonWidth     =   2249
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CLIENTES"
            Description     =   "BOT_CLIENTES"
            Object.ToolTipText     =   "Cadastro de Clientes"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MOTOQUEIROS"
            Description     =   "Bot_Motoqueiros"
            Object.ToolTipText     =   "Cadastro de Motoqueiros"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MOTOCICLETAS"
            Description     =   "MOTOCICLETAS"
            Object.ToolTipText     =   "Cadastro de Motocicletas"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OS"
            Object.ToolTipText     =   "Lançamento de Ordem de Serviços"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SAIR"
            Description     =   "EXIT"
            Object.ToolTipText     =   "Sair do Sistema"
         EndProperty
      EndProperty
      MousePointer    =   10
   End
   Begin VB.Menu men_Aqruivoi 
      Caption         =   "&Arquivo"
      Begin VB.Menu men_Cadastro 
         Caption         =   "Efetuar &Cadastro"
         Begin VB.Menu men_Clientes 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu men_horas 
            Caption         =   "&Horas"
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
Private Sub mem_conexao_Click()
frm_conexao.Show 1
End Sub

Private Sub men_Clientes_Click()
FRM_CadastrodeClientes.Show

End Sub

Private Sub men_horas_Click()

frm_cadHoras.Show

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

Case 5:
    men_Sair_Click

End Select






End Sub
