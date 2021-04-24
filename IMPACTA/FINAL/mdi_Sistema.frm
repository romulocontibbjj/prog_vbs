VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_Sistema 
   BackColor       =   &H8000000C&
   Caption         =   "SISTEMA - FINAL"
   ClientHeight    =   7845
   ClientLeft      =   1170
   ClientTop       =   2220
   ClientWidth     =   11025
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdi_Sistema.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList iml_fotos 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_Sistema.frx":A198
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_Sistema.frx":A5EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbr_sistema 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "iml_fotos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Funcionários"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Produtos"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_funcionarios 
      Caption         =   "&Funcionários"
   End
   Begin VB.Menu mnu_produtos 
      Caption         =   "&Produtos"
   End
End
Attribute VB_Name = "mdi_Sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_funcionarios_Click()
frm_final.Show


End Sub

Private Sub mnu_produtos_Click()
Frm_Produtos.Show

End Sub

Private Sub tbr_sistema_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    frm_final.Show
Case 2
    Frm_Produtos.Show
End Select



End Sub
