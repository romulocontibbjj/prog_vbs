VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{1FB72889-F162-4C84-A0E5-540357575561}#1.0#0"; "xpButtonCtl.ocx"
Begin VB.Form frmSendmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email"
   ClientHeight    =   6780
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendMailExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xpButtonCtl.xpButton cmdIPSetup 
      Height          =   345
      Left            =   1470
      TabIndex        =   46
      Top             =   15
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      TX              =   "IP Setup"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmSendMailExample.frx":030A
   End
   Begin xpButtonCtl.xpButton cmdEmail 
      Height          =   345
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      TX              =   "&Email"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmSendMailExample.frx":0326
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   60
      TabIndex        =   23
      Top             =   270
      Width           =   7845
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   390
         Left            =   735
         TabIndex        =   48
         Top             =   5970
         Width           =   855
      End
      Begin RichTextLib.RichTextBox txtMsg 
         Height          =   1365
         Left            =   1710
         TabIndex        =   47
         Top             =   3285
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   2408
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmSendMailExample.frx":0342
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7080
         Top             =   5640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin xpButtonCtl.xpButton cmdBrowse 
         Height          =   330
         Left            =   6000
         TabIndex        =   19
         Top             =   4815
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         TX              =   "&Browse..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmSendMailExample.frx":03BE
      End
      Begin xpButtonCtl.xpButton cmdExit 
         Height          =   350
         Left            =   6210
         TabIndex        =   22
         Top             =   1245
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   609
         TX              =   "&Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmSendMailExample.frx":03DA
      End
      Begin xpButtonCtl.xpButton cmdReset 
         Height          =   350
         Left            =   6210
         TabIndex        =   21
         Top             =   742
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   609
         TX              =   "&Reset"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmSendMailExample.frx":03F6
      End
      Begin xpButtonCtl.xpButton cmdSend 
         Height          =   350
         Left            =   6210
         TabIndex        =   20
         Top             =   240
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   609
         TX              =   "&Send"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmSendMailExample.frx":0412
      End
      Begin VB.TextBox txtBcc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   7
         Top             =   2430
         Width           =   4200
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   2445
         Left            =   6180
         TabIndex        =   25
         Top             =   1725
         Visible         =   0   'False
         Width           =   1470
         Begin VB.CheckBox ckPopLogin 
            Caption         =   "POP Login"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Use Login Authorization When Connecting to a Host"
            Top             =   2100
            Width           =   1260
         End
         Begin VB.CheckBox ckReceipt 
            Caption         =   "Receipt"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Request a Return Receipt"
            Top             =   1510
            Width           =   1035
         End
         Begin VB.ComboBox cboPriority 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Text            =   "cboPriority"
            ToolTipText     =   "Sets the Prioirty of the Mail Message"
            Top             =   840
            Width           =   1055
         End
         Begin VB.CheckBox ckHtml 
            Caption         =   "Html"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Mail Body is HTML / Plain Text"
            Top             =   1260
            Width           =   1035
         End
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   30
            Top             =   3180
            Width           =   1055
         End
         Begin VB.TextBox txtUserName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   2640
            Width           =   1055
         End
         Begin VB.CheckBox ckLogin 
            Caption         =   "Login"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "Use Login Authorization When Connecting to a Host"
            Top             =   1800
            Width           =   915
         End
         Begin VB.OptionButton optEncodeType 
            Caption         =   "MIME"
            Height          =   195
            Index           =   0
            Left            =   110
            TabIndex        =   27
            ToolTipText     =   "Use MIME encoding for Mail & Attachments."
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optEncodeType 
            Caption         =   "UUEncode"
            Height          =   195
            Index           =   1
            Left            =   110
            TabIndex        =   26
            ToolTipText     =   "Use UU Encoding for Attachments."
            Top             =   540
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   2460
            Width           =   975
         End
      End
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   1740
         TabIndex        =   24
         Top             =   5250
         Width           =   4200
      End
      Begin VB.TextBox txtCcName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   5
         Top             =   1710
         Width           =   4200
      End
      Begin VB.TextBox txtCc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   6
         Top             =   2070
         Width           =   4200
      End
      Begin VB.TextBox txtAttach 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4830
         Width           =   4200
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   8
         Top             =   2790
         Width           =   4200
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   2
         Top             =   630
         Width           =   4200
      End
      Begin VB.TextBox txtFromName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   1
         Top             =   270
         Width           =   4200
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   4
         Top             =   1350
         Width           =   4200
      End
      Begin VB.TextBox txtToName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   3
         Top             =   990
         Width           =   4200
      End
      Begin MSComDlg.CommonDialog cmDialog 
         Left            =   480
         Top             =   3990
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblBcc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bcc: Email"
         Height          =   195
         Left            =   555
         TabIndex        =   9
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   915
         TabIndex        =   10
         Top             =   5310
         Width           =   540
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progress  "
         Height          =   195
         Left            =   3450
         TabIndex        =   11
         Top             =   6135
         Width           =   870
      End
      Begin VB.Label lblCcName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc: Name"
         Height          =   195
         Left            =   600
         TabIndex        =   45
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label lblCC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc: Email"
         Height          =   195
         Left            =   630
         TabIndex        =   44
         Top             =   2085
         Width           =   825
      End
      Begin VB.Label lblAttach 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment"
         Height          =   195
         Left            =   480
         TabIndex        =   43
         Top             =   4890
         Width           =   975
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         Height          =   195
         Left            =   720
         TabIndex        =   42
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   195
         Left            =   810
         TabIndex        =   41
         Top             =   2820
         Width           =   645
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender Email"
         Height          =   195
         Left            =   315
         TabIndex        =   40
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender Name"
         Height          =   195
         Left            =   285
         TabIndex        =   39
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient Email"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Label lblToName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient Name"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1020
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   60
      TabIndex        =   12
      Top             =   270
      Width           =   7845
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Server Setup"
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   720
         TabIndex        =   13
         Top             =   2205
         Width           =   6120
         Begin VB.TextBox txtPopServer 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1770
            MaxLength       =   20
            TabIndex        =   15
            Top             =   885
            Width           =   3885
         End
         Begin VB.TextBox txtServer 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1770
            MaxLength       =   20
            TabIndex        =   14
            Top             =   450
            Width           =   3885
         End
         Begin VB.Label lblPopServer 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "POP3 Server"
            Height          =   195
            Left            =   390
            TabIndex        =   16
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label lblServer 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Server"
            Height          =   195
            Left            =   375
            TabIndex        =   17
            Top             =   480
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "frmSendmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub InsertPictureInRichTextBox(RTB As RichTextBox, Picture As StdPicture)
    ' copy into the clipboard
    ' Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData Picture
    ' paste into the RichTextBox control
    SendMessage RTB.hwnd, WM_PASTE, 0, 0
End Sub



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEmail_Click()
Frame2.Visible = True
Frame1.Visible = False

End Sub

Private Sub cmdIPSetup_Click()
Frame1.Visible = True
Frame2.Visible = False


End Sub

Private Sub cmdSave_Click()
Dim rec As New ADODB.Recordset
If Trim(txtServer.Text) = "" Then
  MsgBox "Enter the Server IP Address or Name", vbInformation + vbOKOnly, "Email Setup"
  txtServer.SetFocus
  Exit Sub
End If
soft.ADOExecute ("Delete * from tblSMTPAddress")
soft.ADOadd_rec "tblSMTPAddress"
soft.ADOSetValue "SMTPServerIP", txtServer.Text
soft.ADOSetValue "POP3ServerIP", txtPopServer.Text
soft.ADOupdate
End Sub

Private Sub Command1_Click()
InsertPictureInRichTextBox txtMsg, LoadPicture("G:\NewEasyAccounts\BalanceSheet10000.jpg")
End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event
    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    lblProgress = ""
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    MsgBox "Send Successful!"
    lblProgress = ""

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub Form_Load()
Dim rec As New ADODB.Recordset
  Dim ret As String
    ' *****************************************************************************
    ' Required to activate the vbSendMail component.
    ' *****************************************************************************
    Me.Top = 1200
    Me.Left = 2000
Me.BackColor = RGB(214, 213, 219)
Frame1.BackColor = RGB(214, 213, 219)
Frame2.BackColor = RGB(214, 213, 219)
Frame3.BackColor = RGB(214, 213, 219)
fraOptions.BackColor = RGB(214, 213, 219)
    Set poSendMail = New clsSendMail

    'With Me
'        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
'        .fraOptions.Height = 2475
'        .lblProgress = ""
'    End With

    cboPriority.AddItem "Normal"
    cboPriority.AddItem "High"
    cboPriority.AddItem "Low"
    cboPriority.ListIndex = 0

Set rec = soft.ADOSearch("select * from tblLogin where UserID=" & UserID)
If Not rec.EOF Then
  txtFromName.Text = "" & rec("Name")
  txtFrom.Text = "" & rec("EmailAddress")
End If
rec.Close
Set rec = soft.ADOSearch("select * from tblSMTPAddress")
If Not rec.EOF Then
  txtServer.Text = "" & rec(0)
  txtPopServer.Text = "" & rec(1)
Else
End If
rec.Close
'    CenterControlsVertical 100, False, txtServer, txtPopServer, txtFromName, txtFrom, txtToName, txtTo, txtCcName, txtCc, txtBcc, txtSubject, txtMsg, txtAttach, lstStatus, lblProgress
'   ' AlignControlsTop False, txtServer, lblServer, cmdSend
'    CenterControlsHorizontal 300, False, lblServer, txtServer, cmdSend
'    AlignControlsLeft False, lblServer, lblPopServer, lblFromName, lblFrom, lblToName, lblTo, lblCcName, lblCC, lblBcc, lblSubject, lblMsg, lstStatus, lblAttach, lblStatus
'
'    CenterControlRelativeVertical lblServer, txtServer
'    CenterControlRelativeVertical lblPopServer, txtPopServer
'    'CenterControlRelativeVertical cmdSend, txtServer
'    CenterControlRelativeVertical lblFromName, txtFromName
'    'CenterControlRelativeVertical cmdReset, txtPopServer
'    CenterControlRelativeVertical lblFrom, txtFrom
'    CenterControlRelativeVertical lblToName, txtToName
'    'CenterControlRelativeVertical cmdExit, txtFrom
'    CenterControlRelativeVertical lblTo, txtTo
'    CenterControlRelativeVertical lblCcName, txtCcName
'    CenterControlRelativeVertical lblCC, txtCc
'    CenterControlRelativeVertical lblBcc, txtBcc
'    CenterControlRelativeVertical lblSubject, txtSubject
'    CenterControlRelativeVertical lblAttach, txtAttach
'    CenterControlRelativeVertical cmdBrowse, txtAttach
'    AlignControlsTop False, txtMsg, lblMsg
'    AlignControlsTop False, lstStatus, lblStatus
'
'    fraOptions.Top = txtTo.Top - 135
'
'    AlignControlsLeft True, txtServer, txtPopServer, txtFromName, txtFrom, txtToName, txtTo, txtCcName, txtCc, txtBcc, txtSubject, txtMsg, lstStatus, txtAttach, lblProgress
'    AlignControlsLeft True, cmdSend, cmdReset, cmdExit, cmdBrowse, fraOptions

  '  lblPopServer.Visible = False
   ' txtPopServer.Visible = False

    Me.Show

    'RetrieveSavedValues

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' *****************************************************************************
    ' Unload the component before quiting.
    ' *****************************************************************************

    Set poSendMail = Nothing

End Sub

Private Sub RetrieveSavedValues()

    ' *****************************************************************************
    ' Retrieve saved values by reading the components 'Persistent' properties
    ' *****************************************************************************
'    poSendMail.PersistentSettings = True
'    txtServer.text = poSendMail.SMTPHost
'    txtPopServer.text = poSendMail.POP3Host
'    txtFrom.text = poSendMail.from
'    txtFromName.text = poSendMail.FromDisplayName
'    txtUserName = poSendMail.Username
'    optEncodeType(poSendMail.EncodeType).Value = True
'    If poSendMail.UseAuthentication Then ckLogin = vbChecked Else ckLogin = vbUnchecked

End Sub

Private Sub cboPriority_Click()

    Select Case cboPriority.ListIndex

        Case 0: etPriority = NORMAL_PRIORITY
        Case 1: etPriority = HIGH_PRIORITY
        Case 2: etPriority = LOW_PRIORITY

    End Select

End Sub

Private Sub cboPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case 38, 40

        Case Else: KeyCode = 0

    End Select

End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ckHtml_Click()

    If ckHtml.Value = vbChecked Then bHtml = True Else bHtml = False

End Sub

Private Sub ckLogin_Click()

    If ckLogin.Value = vbChecked Then
        bAuthLogin = True
        fraOptions.Height = 3555
    Else
        bAuthLogin = False
        If ckPopLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckPopLogin_Click()

    If ckPopLogin.Value = vbChecked Then
        bPopLogin = True
        lblPopServer.Visible = True
        txtPopServer.Visible = True
        fraOptions.Height = 3555
    Else
        bPopLogin = False
        lblPopServer.Visible = False
        txtPopServer.Visible = False
        If ckLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckReceipt_Click()

    If ckReceipt.Value = vbChecked Then bReceipt = True Else bReceipt = False

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = ""
    RetrieveSavedValues

End Sub

Private Sub AlignControlsLeft(StandardizeWidth As Boolean, base As Object, ParamArray cnts())

    
    On Error Resume Next

    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Left = base.Left
        If StandardizeWidth Then cnts(i).Width = base.Width
    Next

End Sub

Private Sub CenterControlsVertical(space As Single, AlignLeft As Boolean, ParamArray cnts())

    

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseLeft As Single

    Dim sngParentHeight As Single

    sngParentHeight = Me.ScaleHeight

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Height
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))
    cnts(0).Top = (sngParentHeight - sngTotalSpace) / 2

    sngBaseLeft = cnts(0).Left

    For i = 1 To UBound(cnts)
        cnts(i).Top = cnts(i - 1).Top + cnts(i - 1).Height + space
        If AlignLeft Then cnts(i).Left = sngBaseLeft
    Next

End Sub

Private Sub CenterControlHorizontal(child As Object)

    child.Left = (Me.ScaleWidth - child.Width) / 2

End Sub

Public Sub CenterControlsHorizontal(space As Single, AlignTop As Boolean, ParamArray cnts())

    

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseTop As Single
    Dim sngParentWidth As Single

    sngParentWidth = Me.ScaleWidth

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Width
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))

    cnts(0).Left = (sngParentWidth - sngTotalSpace) / 2
    sngBaseTop = cnts(0).Top

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        If AlignTop Then cnts(i).Top = sngBaseTop
    Next

End Sub

Public Sub AlignControlsTop(StandardizeHeight As Boolean, base As Object, ParamArray cnts())


    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Top = base.Top
        If StandardizeHeight Then cnts(i).Height = base.Height
    Next

End Sub

Public Sub CenterControlRelativeVertical(ctl As Object, RelativeTo As Object)

    On Error Resume Next
    ctl.Top = RelativeTo.Top + ((RelativeTo.Height - ctl.Height) / 2)

End Sub

Public Sub SetHorizontalDistance(distance As Single, StandardizeWidth As Boolean, AlignTop As Boolean, ParamArray cnts())
    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(cnts)
        If StandardizeWidth Then cnts(i).Width = cnts(i - 1).Width
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + distance
        If AlignTop Then cnts(i).Top = cnts(i - 1).Top
    Next

End Sub

Public Sub CenterControlsRelativeHorizontal(RelativeTo As Object, space As Single, ParamArray cnts())


    On Error Resume Next
    Dim sngTotalWidth As Single
    Dim i As Integer
    For i = 0 To UBound(cnts)
        sngTotalWidth = sngTotalWidth + cnts(i).Width
        If i < UBound(cnts) Then sngTotalWidth = sngTotalWidth + space
    Next

    cnts(0).Left = RelativeTo.Left + ((RelativeTo.Width - sngTotalWidth) / 2)

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        cnts(i).Top = cnts(0).Top
    Next

End Sub

Public Sub ClearTextBoxesOnForm()


    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next

End Sub


Private Sub SSTab1_DblClick()

End Sub

Private Sub txtPopServer_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 13) Then
        KeyAscii = 0
  End If
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 46 And InStr(1, txtServer.Text, ".") > 0 Then
'        KeyAscii = 0
'  End If
  If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 13) Then
        KeyAscii = 0
  End If
End Sub
