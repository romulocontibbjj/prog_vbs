VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form5 
   BackColor       =   &H80000000&
   Caption         =   "Form5"
   ClientHeight    =   1530
   ClientLeft      =   5895
   ClientTop       =   2760
   ClientWidth     =   4155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   1530
   ScaleWidth      =   4155
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog WinDia 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xfile As String

WinDia.DialogTitle = "TESTE - ROMULO"
WinDia.ShowOpen

xfile = WinDia.FileName

MsgBox xfile
End Sub

Private Sub Command2_Click()
WinDia.DialogTitle = "TESTE - SALVAR"
WinDia.ShowSave

End Sub

Private Sub Form_KeyDown(KeyAsc As Integer, Shift As Integer)

If KeyAsc = 27 Then
    Unload Me
End If



End Sub



