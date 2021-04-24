VERSION 5.00
Object = "{76803B87-EB3A-11D1-9BD3-444553540000}#1.0#0"; "FRMAGIC.OCX"
Begin VB.Form FRMTray 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RomSoft System Tray Icon"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Formulário_Mágico_RomSoft.SystemTray SystemTray1 
      Left            =   3120
      Top             =   120
      _ExtentX        =   2223
      _ExtentY        =   847
      TrayIcon        =   "FRMTray.frx":0000
   End
   Begin VB.CheckBox CHKShowIcon 
      Caption         =   "Exibir Ícone"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.PictureBox PICElement 
      Height          =   550
      Index           =   5
      Left            =   3600
      Picture         =   "FRMTray.frx":0452
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      ToolTipText     =   "Click Here To Change The TrayIcon"
      Top             =   2160
      Width           =   550
   End
   Begin VB.PictureBox PICElement 
      Height          =   550
      Index           =   4
      Left            =   3000
      Picture         =   "FRMTray.frx":0894
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      ToolTipText     =   "Click Here To Change The TrayIcon"
      Top             =   2160
      Width           =   550
   End
   Begin VB.PictureBox PICElement 
      Height          =   550
      Index           =   3
      Left            =   2400
      Picture         =   "FRMTray.frx":0CD6
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      ToolTipText     =   "Click Here To Change The TrayIcon"
      Top             =   2160
      Width           =   550
   End
   Begin VB.PictureBox PICElement 
      Height          =   550
      Index           =   2
      Left            =   1800
      Picture         =   "FRMTray.frx":1118
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      ToolTipText     =   "Click Here To Change The TrayIcon"
      Top             =   2160
      Width           =   550
   End
   Begin VB.PictureBox PICElement 
      Height          =   550
      Index           =   1
      Left            =   1200
      Picture         =   "FRMTray.frx":155A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      ToolTipText     =   "Click Here To Change The TrayIcon"
      Top             =   2160
      Width           =   550
   End
   Begin VB.PictureBox PICElement 
      Height          =   550
      Index           =   0
      Left            =   600
      Picture         =   "FRMTray.frx":199C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      ToolTipText     =   "Click Here To Change The TrayIcon"
      Top             =   2160
      Width           =   550
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mudar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1380
      Width           =   975
   End
   Begin VB.TextBox TXTTip 
      Height          =   285
      Left            =   600
      MaxLength       =   60
      TabIndex        =   3
      Top             =   1425
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Para Sair, Clique com o botão Direito sobre o ícone e selecione SAIR"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Texto"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem (Evento TrayAction)"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2235
   End
   Begin VB.Label LBLMessage 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
   Begin VB.Menu MNUPopUp 
      Caption         =   "MenuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnurest 
         Caption         =   "&Restaurar"
      End
      Begin VB.Menu mnumin 
         Caption         =   "Minimi&zar"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu MNURemove 
         Caption         =   "Re&mover este Ícone"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "FRMTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHKShowIcon_Click()
If CHKShowIcon.Value = 1 Then
    SystemTray1.AdicionaIcone  'Show the icon
Else
    SystemTray1.RemoveIcone 'Delete the icon
End If
End Sub

Private Sub Command1_Click()
SystemTray1.TextoDica = TXTTip.Text 'Change the tip
End Sub

Private Sub Form_Load()
SystemTray1.AdicionaIcone  'Cria o ícone
End Sub

Private Sub MnuExit_Click()
SystemTray1.RemoveIcone 'Deleta o Ícone
End
End Sub



Private Sub mnumin_Click()
FRMTray.WindowState = 1
End Sub

Private Sub MNURemove_Click()
SystemTray1.RemoveIcone 'Delete the TrayIcon
CHKShowIcon.Value = 0
End Sub

Private Sub mnurest_Click()
FRMTray.WindowState = 0
End Sub

Private Sub PICElement_Click(Index As Integer)
SystemTray1.TrocaIcone PICElement(Index).Picture  'Muda o Ícone

End Sub

Private Sub SystemTray1_TrayAction(Action As Integer)
'Esse evento ocorre quando
Select Case Action
    Case 0 'Mouse Move
        Beep
    Case 1 'Left mouse down
        LBLMessage.Caption = "Você pressionou o botão esquerdo"
    Case 2 'Left Mouse UP
        LBLMessage.Caption = "Você soltou o botão esquerdo"
    Case 3 'left double click
        MsgBox "Você clicou duas vezes"
        Me.SetFocus
    Case 4 'Right mouse down
        LBLMessage.Caption = "Você pressionou o botão direito"
    Case 5 'Right mouse Up
        PopupMenu MNUPopUp
        LBLMessage.Caption = "Você soltou o botão direito"
    Case 6 'Right Double Click
        MsgBox "Mocê clicou duas vezes com o botão direito"
End Select
        
        
End Sub
