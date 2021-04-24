VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frm_inicio 
   Caption         =   "Form2"
   ClientHeight    =   6435
   ClientLeft      =   1215
   ClientTop       =   1545
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   ScaleHeight     =   6435
   ScaleWidth      =   8490
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frm_inicio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "img_fotos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "drv_dir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fil_dir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dir_dir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chk_automatica"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tmr_fotos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Timer tmr_fotos 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1560
         Top             =   5280
      End
      Begin VB.CheckBox chk_automatica 
         Caption         =   "Automático"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   5280
         Width           =   1095
      End
      Begin VB.DirListBox dir_dir 
         Height          =   2115
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
      End
      Begin VB.FileListBox fil_dir 
         Height          =   1845
         Left            =   120
         Pattern         =   "*.jpg; *.bmp"
         TabIndex        =   2
         Top             =   3360
         Width           =   2775
      End
      Begin VB.DriveListBox drv_dir 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image img_fotos 
         BorderStyle     =   1  'Fixed Single
         Height          =   5055
         Left            =   3240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frm_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk_automatica_Click()
tmr_fotos.Enabled = Not tmr_fotos.Enabled
    
End Sub

Private Sub dir_dir_Change()
fil_dir.Path = dir_dir.Path

End Sub

Private Sub drv_dir_Change()
On Error GoTo RRC
dir_dir.Path = drv_dir.Drive
RRC:
If Err.Number = 68 Then
    MsgBox "Unidade Inválida", vbCritical, UCase(drv_dir) & "\"
End If


End Sub

Private Sub fil_dir_Click()
img_fotos.Picture = LoadPicture(fil_dir.Path & "\" & fil_dir.FileName)

End Sub

Private Sub Form_Load()
drv_dir.Drive = "D:\"
dir_dir.Path = drv_dir.Drive
End Sub

Private Sub img_fotos_Click()
frm_imagem.Show
frm_imagem.img_ampliada.Picture = LoadPicture(fil_dir.Path & "\" & fil_dir.FileName)
frm_imagem.Caption = UCase(fil_dir.FileName)
DoEvents

End Sub

Private Sub tmr_fotos_Timer()
Dim x As Integer
x = 0

img_fotos.Picture = LoadPicture(fil_dir.Path & fil_dir.List(x))

x = x - 1

If x = fil_dir.ListCount Then x = 0





    

End Sub











