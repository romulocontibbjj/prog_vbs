VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Message Format"
      Height          =   2925
      Left            =   30
      TabIndex        =   5
      Top             =   45
      Width           =   6525
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   325
         Left            =   5280
         TabIndex        =   4
         Top             =   2460
         Width           =   1065
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   325
         Left            =   3930
         TabIndex        =   3
         Top             =   2460
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         Caption         =   "Advanced Settings"
         Height          =   1110
         Left            =   510
         TabIndex        =   8
         Top             =   1260
         Width           =   5775
         Begin VB.OptionButton optEnCode 
            Caption         =   "Unencode"
            Height          =   285
            Left            =   600
            TabIndex        =   2
            Top             =   705
            Width           =   2475
         End
         Begin VB.OptionButton optMIME 
            Caption         =   "MIME"
            Height          =   285
            Left            =   600
            TabIndex        =   1
            Top             =   345
            Value           =   -1  'True
            Width           =   2475
         End
      End
      Begin VB.ComboBox cmbMessageFormat 
         Height          =   315
         ItemData        =   "frmOptions.frx":08CA
         Left            =   3270
         List            =   "frmOptions.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   795
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send in this message format"
         Height          =   195
         Left            =   525
         TabIndex        =   7
         Top             =   840
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose a format for outgoing mail and change advanced settings"
         Height          =   195
         Left            =   495
         TabIndex        =   6
         Top             =   405
         Width           =   5610
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbMessageFormat_Click()
    If cmbMessageFormat.ListIndex = 1 Then
        MyEncodeType = MIME_ENCODE
        bHtml = True
    Else
        MyEncodeType = UU_ENCODE
        bHtml = False
    End If
End Sub

Private Sub cmdApply_Click()
Dim ret As Boolean
If cmbMessageFormat.ListIndex = 1 And optMIME = True Then
   ret = MAILFORMAT(True, MIME_ENCODE)
   MyEncodeType = MIME_ENCODE
   bHtml = True
End If
If cmbMessageFormat.ListIndex = 1 And optMIME = False Then
   ret = MAILFORMAT(True, UU_ENCODE)
   MyEncodeType = UU_ENCODE
   bHtml = True
End If

If cmbMessageFormat.ListIndex = 0 And optMIME = True Then
   ret = MAILFORMAT(False, MIME_ENCODE)
   bHtml = False
   MyEncodeType = MIME_ENCODE
End If
If cmbMessageFormat.ListIndex = 0 And optMIME = False Then
   ret = MAILFORMAT(False, UU_ENCODE)
   bHtml = False
   MyEncodeType = UU_ENCODE
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim ret As String
   ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "MAILFORMAT")
   If ret <> "" Then
        If Split(ret, "%")(1) = "0" Then
            optMIME.Value = True
        Else
            optEnCode.Value = True
        End If
        If Split(ret, "%")(0) = True Then
            cmbMessageFormat.ListIndex = 1
        Else
            cmbMessageFormat.ListIndex = 0
        End If
    End If
End Sub

Private Sub optEnCode_Click()
MyEncodeType = UU_ENCODE
End Sub

Private Sub optMIME_Click()
MyEncodeType = MIME_ENCODE
End Sub
