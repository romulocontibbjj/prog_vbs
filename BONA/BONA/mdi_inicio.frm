VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_inicio 
   BackColor       =   &H8000000C&
   Caption         =   "ROMULO"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   11700
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdi_inicio.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_inicio.frx":11353
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_inicio.frx":125D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_inicio.frx":13857
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_inicio.frx":14AD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_inicio.frx":15D5B
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
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1111
      ButtonWidth     =   1799
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "BONAGURA"
            Description     =   "BONA"
            Object.ToolTipText     =   "Compara Arquivos Bonagura"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DIVERSOS"
            Description     =   "Diversos"
            Object.ToolTipText     =   "Vários Tipos de Serviços"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "FOX"
            Description     =   "FOX"
            Object.ToolTipText     =   "Especial FOX FILMES"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DIVERSOS II"
            Description     =   "DIVERSOS II"
            Object.ToolTipText     =   "Continuação de Diversos"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SAIR"
            Description     =   "SAIR"
            Object.ToolTipText     =   "QUERO SAIR DAQUI!!!!!"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Menu mem_programa 
      Caption         =   "&Programas"
      Begin VB.Menu mem_exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "mdi_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mem_exit_Click()
Unload Me

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
frm_bonagura.Show
Case 2
frm_diversos.Show
Case 3
frm_fox.Show
Case 4
frm_diversosII.Show
Case 5
Unload Me

End Select


End Sub
