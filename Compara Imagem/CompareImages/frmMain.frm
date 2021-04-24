VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare Images"
   ClientHeight    =   3120
   ClientLeft      =   5085
   ClientTop       =   2775
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5160
   Begin MSComDlg.CommonDialog objCD 
      Left            =   1320
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "&Compare"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Load..."
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   " Load Image "
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Load..."
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   " Load Image "
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   1
      Left            =   2640
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":335F
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Main image comparing object.
Private objCompare As clsImgComp

Private Sub cmdBrowse_Click(Index As Integer)
BrowseForPicture Index
End Sub

Private Sub BrowseForPicture(Index As Integer)
On Error GoTo ErrorHandler

With objCD
    .DialogTitle = "Open Image #" & Index + 1
    .Filter = "Image Files|*.bmp;*.jpg;*.gif;*.rle;*.tiff;*.jpeg;*.tif;*.dib"
    .ShowOpen
    
    If Len(.FileName) > 0 Then
        Set picBox(Index).Picture = LoadPicture(.FileName, , vbLPDefault)
    End If

End With

Exit Sub

ErrorHandler:

    MsgBox Err.Description, vbCritical, "Error #" & Err.Number
End Sub

Private Sub cmdCompare_Click()
Dim bolAlike As Boolean

bolAlike = objCompare.CompareImage(picBox(0).hdc, picBox(0).Image.Handle, picBox(1).hdc, picBox(1).Image.Handle)

If bolAlike = True Then
    MsgBox "As Imagens São Iguais", vbInformation, "Images Equal"
Else
    MsgBox "As Imagens São Diferentes", vbInformation, "Images Different"
End If
End Sub

Private Sub Form_Load()
'Create a new instance of the image comparing object.
Set objCompare = New clsImgComp
End Sub
