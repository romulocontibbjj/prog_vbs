VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Email Test Utility"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkArray 
      Caption         =   "Email Client Running"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdlgAttachment 
      Left            =   3840
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton butArray 
      Caption         =   "..."
      Height          =   300
      Index           =   3
      Left            =   4160
      TabIndex        =   8
      Top             =   4200
      Width           =   300
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Text            =   "ReadMe.txt"
      Top             =   4200
      Width           =   3855
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "Close resources after sends"
      Height          =   615
      Index           =   1
      Left            =   3120
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame fraArray 
      Caption         =   "Send Type"
      Height          =   1095
      Index           =   0
      Left            =   1680
      TabIndex        =   23
      Top             =   4680
      Width           =   1335
      Begin VB.OptionButton optArray 
         Caption         =   "Use CDO"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optArray 
         Caption         =   "Use MAPI"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optArray 
         Caption         =   "Use Outlook"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox txtArray 
      Height          =   1005
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "Email List Enabled"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton butArray 
      Caption         =   "Save List Info"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   5970
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7594
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton butArray 
      Caption         =   "Get List Info"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Text            =   "DistList"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton butArray 
      Caption         =   "Send Mail"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtArray 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblArray 
      Caption         =   "Attachment (clear name for no attachment):"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lblArray 
      Caption         =   "Body Text:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblArray 
      Caption         =   "Subject:"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblArray 
      Caption         =   "(Low,Normal,High)"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   5460
      Width           =   1695
   End
   Begin VB.Label lblArray 
      Caption         =   "Importance:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblArray 
      Caption         =   "Email List Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblArray 
      Caption         =   "Send To (semicolon-delimited list):"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==========================================================================
'Copyright © 2001 by Stan Schultes, All Rights Reserved
'
'   frmMain.frm
'   Startup form for EmailTest
'   EmailTest - send message utility
'   Date Created: 11-Nov-2000
'
'   Notes:
'   The EmailTest sample application demonstrates how to send email from a
'   VB application. Email settings are stored in the registry by EmailList,
'   allowing the application to send to several lists with different settings.
'   See comments in CEmailReg.cls for details.
'
'   In practice, you would need one of the 3 email class types - CEmailCDO,
'   CEmailOL, or goMailMAPI (a form that acts like a class for this application).
'   See the article text for choosing which might fit your needs the best. You
'   also need the CEmailReg.cls module, and the EmailTest.bas module.
'
'   In a production application, you might consider using a setup form similar
'   to this frmMain.frm module to allow configuration of the Email settings.
'   You would do this to keep users of your application from having to edit the
'   Registry directly.
'
'- in CheckAttachment, test attachment existence
'==========================================================================
'
'Form Controls Constant Definitions
'-- Text Boxes
Private Const kTxtEmailList As Long = 0
Private Const kTxtSendTo As Long = 1
Private Const kTxtImportance As Long = 2
Private Const kTxtSubject As Long = 3
Private Const kTxtBody As Long = 4
Private Const kTxtAttachment As Long = 5
'-- Option Buttons
Private Const kOptCDO As Long = 0
Private Const kOptMAPI As Long = 1
Private Const kOptOutlook As Long = 2
'-- Command Buttons
Private Const kButSend As Long = 0
Private Const kButGet As Long = 1
Private Const kButSave As Long = 2
Private Const kButChoose As Long = 3
'-- Check Boxes
Private Const kChkEnabled As Long = 0
Private Const kChkCloseRes As Long = 1
Private Const kChkClientActive As Long = 2

Private Sub butArray_Click(Index As Integer)
'handles form button clicks
Dim sMsg As String
    goReg.EmailList = txtArray(kTxtEmailList)   'set listname
    Select Case Index
    Case kButSend
        Screen.MousePointer = vbHourglass
        StatusText "Sending..."
        sMsg = SendMail()
        StatusText sMsg
        Screen.MousePointer = vbDefault
    Case kButGet
        GetEmailList
        StatusText "Choose a message type & attachment and click Send Mail"
    Case kButSave
        SaveEmailList
        StatusText "Configuration saved for EmailList: " & txtArray(kTxtEmailList)
    Case kButChoose
        ChooseAttachment
    Case Else
        StatusText "Unrecognized button: " & CStr(Index)
    End Select
End Sub

Private Sub optArray_Click(Index As Integer)
'handles option button clicks
'here set Subject
Dim sMsg As String
    Select Case Index
    Case kOptCDO
        sMsg = "Test CDO Message"
    Case kOptMAPI
        sMsg = "Test MAPI Message"
    Case kOptOutlook
        sMsg = "Test Outlook Message"
    Case Else
        sMsg = "Unknown Send Type"
    End Select
    txtArray(kTxtSubject) = sMsg
End Sub

Private Sub Form_Load()
'app startup
    Set goReg = New CEmailReg   'distribution list manager
    CheckInit                   'set default list if first run
    chkArray(kChkClientActive).Value = vbChecked    'assume email client running
    StatusText "Enter the Email List and click Get List Info"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'release resources that may be in use
    Unload goMailMAPI       'MAPI class is a form
    Set goMailCDO = Nothing
    Set goMailOL = Nothing
    Set goReg = Nothing
End Sub

Private Sub StatusText(ByVal StatusValue As String)
'write to the status bar
    sbrStatus.Panels(1).Text = Trim$(StatusValue)
    DoEvents            'let the status draw
End Sub

Public Property Get CloseResources() As Boolean
'returns value of Close Resources CheckBox (Read-only form property)
    CloseResources = CBool(chkArray(kChkCloseRes).Value)
End Property

Public Sub CheckInit()
'sets up initial EmailList if first run
    goReg.EmailList = "DistList"
    If goReg.SendTo = "" Then
        goReg.Importance = "Normal"
        goReg.SendTo = "Enter your email address here and click Save List Info"
        goReg.Enabled = True
    End If
End Sub

Public Function SendMail() As String
'sends mail based on form choices, returns status message
Dim lStatus As Long
Dim oTimer As New CElapsedTime      'used to time operations
    If optArray(kOptCDO) Then
        oTimer.StartTheClock
        lStatus = goMailCDO.Send(txtArray(kTxtEmailList), CBool(chkArray(kChkClientActive).Value), txtArray(kTxtSubject), txtArray(kTxtBody), txtArray(kTxtAttachment))
        oTimer.StopTheClock
        If lStatus = 0 Then
            SendMail = "CDO Send Succeeded, elapsed: " & oTimer.Elapsed & " ms"
        Else
            SendMail = "CDO Send Failed, Err: " & lStatus & ", elapsed: " & oTimer.Elapsed & " ms"
        End If
        If CloseResources Then Set goMailCDO = Nothing
    ElseIf optArray(kOptMAPI) Then
        'MAPI version uses a form to hold the MAPI controls
        oTimer.StartTheClock
        lStatus = goMailMAPI.Send(txtArray(kTxtEmailList), CBool(chkArray(kChkClientActive).Value), txtArray(kTxtSubject), txtArray(kTxtBody), txtArray(kTxtAttachment))
        oTimer.StopTheClock
        If lStatus = 0 Then
            SendMail = "MAPI Send Succeeded, elapsed: " & oTimer.Elapsed & " ms"
        Else
            SendMail = "MAPI Send Failed, Err: " & lStatus & ", elapsed: " & oTimer.Elapsed & " ms"
        End If
        If CloseResources Then Unload goMailMAPI
    ElseIf optArray(kOptOutlook) Then
        oTimer.StartTheClock
        lStatus = goMailOL.Send(txtArray(kTxtEmailList), CBool(chkArray(kChkClientActive).Value), txtArray(kTxtSubject), txtArray(kTxtBody), txtArray(kTxtAttachment))
        oTimer.StopTheClock
        If lStatus = 0 Then
            SendMail = "OL Send Succeeded, elapsed: " & oTimer.Elapsed & " ms"
        Else
            SendMail = "OL Send Failed, Err: " & lStatus & ", elapsed: " & oTimer.Elapsed & " ms"
        End If
        If CloseResources Then Set goMailOL = Nothing
    End If
End Function

Public Sub GetEmailList()
'get EmailList properties to form
    'assumes goReg.EmailList has already been set
    txtArray(kTxtSendTo) = goReg.SendTo
    txtArray(kTxtImportance) = goReg.Importance
    chkArray(kChkEnabled).Value = IIf(goReg.Enabled, vbChecked, vbUnchecked)
    'put other defaults on the form
    If CBool(optArray(kOptCDO)) Then
        txtArray(kTxtSubject) = "Test CDO Message"
    ElseIf CBool(optArray(kOptMAPI)) Then
        txtArray(kTxtSubject) = "Test MAPI Message"
    ElseIf CBool(optArray(kOptOutlook)) Then
        txtArray(kTxtSubject) = "Test Outlook Message"
    End If
    txtArray(kTxtBody) = "Message Body Text"
End Sub

Public Sub SaveEmailList()
'save EmailList properties from form
    'assumes goReg.EmailList has already been set
    goReg.SendTo = txtArray(kTxtSendTo)
    goReg.Importance = txtArray(kTxtImportance)
    goReg.Enabled = CBool(chkArray(kChkEnabled).Value)
End Sub

Public Sub ChooseAttachment()
'show File Open common dialog
    cdlgAttachment.InitDir = App.Path
    cdlgAttachment.FileName = txtArray(kTxtAttachment)
    cdlgAttachment.ShowOpen
    txtArray(kTxtAttachment) = cdlgAttachment.FileName
End Sub
