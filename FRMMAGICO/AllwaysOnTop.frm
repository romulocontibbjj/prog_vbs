VERSION 5.00
Object = "{76803B87-EB3A-11D1-9BD3-444553540000}#1.0#0"; "FRMAGIC.OCX"
Begin VB.Form FRMAlwaysOnTop 
   Caption         =   "Formulário Sempre Visível RomSoft"
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin Formulário_Mágico_RomSoft.SempreVisivel SempreVisivel1 
      Left            =   120
      Top             =   360
      _ExtentX        =   2223
      _ExtentY        =   847
   End
   Begin VB.CommandButton CMDEnd 
      Caption         =   "&Fim"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox CHKOnTop 
      Caption         =   "Sempre Visível"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "FRMAlwaysOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHKOnTop_Click()
If CHKOnTop.Value = 1 Then
    SempreVisivel1.Enabled = True 'Ativa
Else
    SempreVisivel1.Enabled = False
End If
End Sub

Private Sub CMDEnd_Click()
End
End Sub

Private Sub Form_Load()

End Sub
