VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About PrnInfo.exe"
   ClientHeight    =   2460
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   5535
   ClipControls    =   0   'False
   Icon            =   "FAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "FAbout.frx":0442
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2460
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4350
      TabIndex        =   0
      Top             =   690
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "FAbout.frx":0888
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   300
      MouseIcon       =   "FAbout.frx":0B92
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1620
      Width           =   660
   End
   Begin VB.Label lblComments 
      BackStyle       =   0  'Transparent
      Caption         =   "A demonstration of many of the Win32 print spooler functions. Be sure to visit the author's site for the most recent updates."
      Height          =   735
      Left            =   300
      TabIndex        =   5
      Top             =   1920
      Width           =   4890
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   180
      X2              =   5400
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   180
      X2              =   5400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   525
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   1350
      Width           =   660
   End
   Begin VB.Label lblActiveX2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Written with Classic VB5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   2985
      TabIndex        =   2
      Top             =   0
      Width           =   2595
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PrnInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   945
      TabIndex        =   1
      Top             =   300
      Width           =   1665
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright (C)1999-2001 Karl E. Peterson, All Rights Reserved
' *********************************************************************
'  Warning: This computer program is protected by copyright law and
'  international treaties. Unauthorized reproduction or distribution
'  of this program, or any portion of it, may result in severe civil
'  and criminal penalties, and will be prosecuted to the maximum
'  extent possible under the law.
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

' API structure definition for Rectangle
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command1_GotFocus()
   Dim r As RECT
   Static BeenThereDoneThat As Boolean

   If Not BeenThereDoneThat Then
      Call GetWindowRect((Command1.hWnd), r)
      Call SetCursorPos(r.Left + (r.Right - r.Left) \ 2, _
                        r.Top + (r.Bottom - r.Top) \ 2)
      BeenThereDoneThat = True
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   '
   ' Setup main form elements
   '
   lblCopyright.Caption = App.LegalCopyright
   lblVersion.Caption = "Version " & _
      Format(App.Major, "#0") & "." & _
      Format(App.Minor, "#00") & "." & _
      Format(App.Revision, "0000")

   lblCompany.Caption = App.CompanyName
   
   Line1.X2 = Me.ScaleWidth - Line1.X1
   Line2.X1 = Line1.X1
   Line2.X2 = Line1.X2
   Line2.Y1 = Line1.Y1 + Screen.TwipsPerPixelY
   Line2.Y2 = Line1.Y2 + Screen.TwipsPerPixelY
   
   Command1.Left = Line1.X2 - Command1.Width
End Sub

Private Sub lblCompany_Click()
   Call ShellExecute(0&, vbNullString, lblCompany.Caption, vbNullString, vbNullString, vbNormalFocus)
End Sub
