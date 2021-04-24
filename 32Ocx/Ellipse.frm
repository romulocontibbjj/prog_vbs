VERSION 5.00
Object = "{76803B87-EB3A-11D1-9BD3-444553540000}#1.0#0"; "FRMAGIC.OCX"
Begin VB.Form Form1 
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin Formulário_Mágico_RomSoft.Elipse Elipse1 
      Left            =   2280
      Top             =   0
      _ExtentX        =   2223
      _ExtentY        =   847
      Largura         =   6165
      Altura          =   3330
      Ajuda           =   ""
   End
   Begin VB.CheckBox CHKElliptical 
      Caption         =   "Elíptico"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton CMDRESTORE 
      Caption         =   "&Restaura"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton CMDMAX 
      Caption         =   "&Maximiza"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton CMDEND 
      Caption         =   "&Fim"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.HScrollBar SCRXini 
      Height          =   255
      LargeChange     =   300
      Left            =   3765
      Max             =   1000
      SmallChange     =   50
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.HScrollBar SCRYIni 
      Height          =   255
      LargeChange     =   300
      Left            =   3765
      Max             =   1000
      SmallChange     =   50
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.HScrollBar SCRHeight 
      Height          =   255
      LargeChange     =   300
      Left            =   3765
      Max             =   20000
      SmallChange     =   50
      TabIndex        =   3
      Top             =   1080
      Value           =   3300
      Width           =   1695
   End
   Begin VB.HScrollBar SCRWIdth 
      Height          =   255
      LargeChange     =   300
      Left            =   3765
      Max             =   20000
      SmallChange     =   50
      TabIndex        =   2
      Top             =   720
      Value           =   6200
      Width           =   1695
   End
   Begin VB.CheckBox CHKFullForm 
      Caption         =   "Sempre Exibir Todo o Form"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton CMDFullForm 
      Caption         =   "Exibir Todo o Form"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "YIni"
      Height          =   195
      Left            =   3435
      TabIndex        =   9
      Top             =   1800
      Width           =   270
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "XIni"
      Height          =   195
      Left            =   3390
      TabIndex        =   8
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Altura"
      Height          =   195
      Left            =   3300
      TabIndex        =   5
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Largura"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CHKElliptical_Click()
If CHKElliptical.Value = 1 Then
    Elipse1.Enabled = True
Else
    Elipse1.Enabled = False
End If
End Sub

Private Sub CHKFullForm_Click()
If CHKFullForm.Value = 1 Then
    CMDFullForm.Enabled = False
Else
    CMDFullForm.Enabled = True
End If
End Sub

Private Sub CMDEND_Click()
End
End Sub

Private Sub CMDFullForm_Click()
    Elipse1.TodoForm = True 'Ajusta a elipse ao formulario
    'As linhas abaixo são usadas apenas para atualizar as barras de rolagem
    SCRWIdth.Value = Elipse1.Largura
    SCRHeight.Value = Elipse1.Altura
    SCRXini.Value = 0
    SCRYIni.Value = 0
End Sub

Private Sub CMDMAX_Click()
Me.WindowState = 2
Elipse1.TodoForm = True
End Sub



Private Sub CMDRESTORE_Click()
Me.WindowState = 0
Elipse1.TodoForm = True
End Sub

Private Sub Form_Resize()
If CHKFullForm.Value = 1 Then
    Elipse1.TodoForm = True 'Ajusta a elipse ao formulario
    'As linhas abaixo são usadas apenas para atualizar as barras de rolagem
    SCRWIdth.Value = Elipse1.Largura
    SCRHeight.Value = Elipse1.Altura
    SCRXini.Value = 0
    SCRYIni.Value = 0
End If
End Sub

Private Sub SCRHeight_Change()
Elipse1.Altura = SCRHeight.Value
End Sub

Private Sub SCRHeight_Scroll()
Elipse1.Altura = SCRHeight.Value
End Sub

Private Sub SCRWIdth_Change()
Elipse1.Largura = SCRWIdth.Value
End Sub

Private Sub SCRWIdth_Scroll()
Elipse1.Largura = SCRWIdth.Value
End Sub

Private Sub SCRXini_Change()
Elipse1.Xini = SCRXini.Value
End Sub

Private Sub SCRXini_Scroll()
Elipse1.Xini = SCRXini.Value
End Sub

Private Sub SCRYIni_Change()
Elipse1.Yini = SCRYIni.Value

End Sub

Private Sub SCRYIni_Scroll()
Elipse1.Yini = SCRYIni.Value
End Sub
