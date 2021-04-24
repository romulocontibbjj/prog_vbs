VERSION 5.00
Begin VB.Form frm_diretor 
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frm_diretor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox fil_Dir 
      Height          =   2430
      Left            =   2520
      Pattern         =   "*.ico;*.bmp;*.jpg"
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.DirListBox Dir 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.DriveListBox Drv_Dir 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lab_Caminho 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Img_fotos 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   4440
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frm_diretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir_Change()
    fil_Dir.Path = Dir.Path
    Img_fotos.Picture = LoadPicture
    lab_Caminho.Caption = fil_Dir.Path
    
    
End Sub

Private Sub Drv_Dir_Change()
On Error GoTo trataUnidade
    Dir.Path = Drv_Dir.Drive
trataUnidade:
    If Err.Number = 68 Then
        MsgBox "Unidade Inválida", vbCritical, UCase(Drv_Dir) & "\"
    End If
    
End Sub

Private Sub fil_Dir_Click()
Img_fotos.Picture = LoadPicture(fil_Dir.Path & "\" & fil_Dir.FileName)
lab_Caminho.Caption = Dir & "\" & fil_Dir & "\" & fil_Dir.FileName

End Sub

Private Sub Form_Load()
Me.Caption = Format(Date, "dddd - dd/mm/yyyy")
lab_Caminho.Caption = fil_Dir.Path
'Dim vSemana As String
'vSemana = InputBox("DIGITE A DATA", "SEMANA")
'vSemana = WeekdayName(Weekday(vSemana))

'MsgBox "Essa Data corresponde a : " & vSemana

End Sub
