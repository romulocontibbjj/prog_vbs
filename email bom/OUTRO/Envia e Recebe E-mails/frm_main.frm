VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frankmailer"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tim_timeout 
      Left            =   5520
      Top             =   5460
   End
   Begin VB.CommandButton cmd_action 
      Caption         =   "Send/Recieve"
      Height          =   555
      Left            =   4560
      TabIndex        =   20
      Top             =   6060
      Width           =   1515
   End
   Begin MSWinsockLib.Winsock wsk_socket 
      Left            =   5040
      Top             =   5460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab tab_mailer 
      Height          =   5895
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SMTP (Outgoing)"
      TabPicture(0)   =   "frm_main.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_msgto"
      Tab(0).Control(1)=   "fra_msgbody"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "POP3 (Incoming)"
      TabPicture(1)   =   "frm_main.frx":2D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_popfolder"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fra_popfolder 
         Caption         =   "Message Folder"
         Height          =   2895
         Left            =   180
         TabIndex        =   18
         Top             =   2760
         Width           =   5655
         Begin VB.FileListBox fil_msgfiles 
            Height          =   2040
            Left            =   720
            Pattern         =   "*.eml"
            TabIndex        =   19
            Top             =   300
            Width           =   4755
         End
         Begin VB.Label lbl_note 
            Caption         =   "Double clicking on the email messages will attempt to open them in your email client."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   435
            Left            =   720
            TabIndex        =   21
            Top             =   2340
            Width           =   4755
         End
         Begin VB.Image img_msgfolder 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":2D32
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Account Details"
         Height          =   2115
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   5655
         Begin VB.TextBox txt_poppassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   17
            Top             =   1500
            Width           =   3675
         End
         Begin Frankmailer.ctl_dynacom dco_popusername 
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   960
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
         End
         Begin Frankmailer.ctl_dynacom dco_popserver 
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   420
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
         End
         Begin VB.Label lbl_poppassword 
            Caption         =   "Password"
            Height          =   195
            Left            =   660
            TabIndex        =   16
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lbl_popusername 
            Caption         =   "Username"
            Height          =   195
            Left            =   660
            TabIndex        =   15
            Top             =   1020
            Width           =   795
         End
         Begin VB.Image img_poppassword 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":303C
            Top             =   1440
            Width           =   480
         End
         Begin VB.Image img_popusername 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":3346
            Top             =   900
            Width           =   480
         End
         Begin VB.Image img_popserver 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":4010
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lbl_popserver 
            Caption         =   "POP3 Server"
            Height          =   255
            Left            =   660
            TabIndex        =   14
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame fra_msgto 
         Caption         =   "Message to / from"
         Height          =   2115
         Left            =   -74820
         TabIndex        =   3
         Top             =   480
         Width           =   5655
         Begin Frankmailer.ctl_dynacom dco_msgto 
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            Top             =   1500
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
         End
         Begin Frankmailer.ctl_dynacom dco_msgfrom 
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   960
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
         End
         Begin Frankmailer.ctl_dynacom dco_smtpserver 
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   420
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
         End
         Begin VB.Label lbl_server 
            Caption         =   "SMTP Server"
            Height          =   255
            Left            =   660
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin VB.Image img_server 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":431A
            Top             =   300
            Width           =   480
         End
         Begin VB.Image img_msgto 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":4624
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image img_msgfrom 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":52EE
            Top             =   840
            Width           =   480
         End
         Begin VB.Label lbl_msgfrom 
            Caption         =   "From"
            Height          =   195
            Left            =   660
            TabIndex        =   5
            Top             =   1020
            Width           =   435
         End
         Begin VB.Label lbl_msgto 
            Caption         =   "To"
            Height          =   195
            Left            =   660
            TabIndex        =   4
            Top             =   1560
            Width           =   435
         End
      End
      Begin VB.Frame fra_msgbody 
         Caption         =   "Email Subject/Body"
         Height          =   2895
         Left            =   -74820
         TabIndex        =   1
         Top             =   2760
         Width           =   5655
         Begin VB.TextBox txt_msgsubject 
            Height          =   315
            Left            =   720
            MaxLength       =   40
            TabIndex        =   6
            Top             =   300
            Width           =   4755
         End
         Begin VB.TextBox txt_msgbody 
            Height          =   1995
            Left            =   720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   4755
         End
         Begin VB.Image img_msgbody 
            Height          =   480
            Left            =   120
            Picture         =   "frm_main.frx":AF0C
            Top             =   300
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************
'SHELL EXECUTE
'**********************************
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Const SW_SHOWNORMAL = 1

'**********************************
'FORM EVENTS
'**********************************
    Private Sub Form_Load()
        'SMTP (OUTGOING)
            With dco_smtpserver
                .Setautocomplete True
                .Setmaxentries 50
                .Loadentries App.Path & "\" & .Name & ".log"
            End With
            With dco_msgto
                .Setautocomplete True
                .Setmaxentries 50
                .Loadentries App.Path & "\" & .Name & ".log"
            End With
            With dco_msgfrom
                .Setautocomplete True
                .Setmaxentries 50
                .Loadentries App.Path & "\" & .Name & ".log"
            End With
        'POP3 INCOMING
            With dco_popserver
                .Setautocomplete True
                .Setmaxentries 50
                .Loadentries App.Path & "\" & .Name & ".log"
            End With
            With dco_popusername
                .Setautocomplete True
                .Setmaxentries 50
                .Loadentries App.Path & "\" & .Name & ".log"
            End With
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        'SMTP (OUTGOING)
            dco_smtpserver.Saveentries App.Path & "\" & "dco_smtpserver.Log"
            dco_msgto.Saveentries App.Path & "\" & "dco_msgto.Log"
            dco_msgfrom.Saveentries App.Path & "\" & "dco_msgfrom.Log"
        'POP3 (INCOMING)
            dco_popserver.Saveentries App.Path & "\" & "dco_popserver.log"
            dco_popusername.Saveentries App.Path & "\" & "dco_popusername.log"
            fil_msgfiles.Path = App.Path
    End Sub

'**********************************
'SEND AND RECIEVE BUTTON
'**********************************
    Private Sub cmd_action_Click()
        Dim socket As Winsock
        Dim timeout As Timer
        Dim messages As Integer
        Set socket = wsk_socket
        Set timeout = tim_timeout
    
        If Not mailbusy Then
            If tab_mailer.Tab = 0 Then
                dco_smtpserver.Addentry dco_smtpserver.Getcurrententry
                dco_msgto.Addentry dco_msgto.Getcurrententry
                dco_msgfrom.Addentry dco_msgfrom.Getcurrententry
                Me.Caption = "Frankmailer - Sending Mail"
                If sendmail(socket, dco_smtpserver.Getcurrententry, 25, dco_msgto.Getcurrententry, dco_msgfrom.Getcurrententry, txt_msgsubject, txt_msgbody, timeout) Then
                    MsgBox "Mail sent", vbOKOnly, "Sending Mail"
                Else
                    MsgBox "Mail not sent", vbOKOnly, "Sending Mail"
                End If
                Unload frm_progress
                frm_main.Caption = "Frankmailer"
            ElseIf tab_mailer.Tab = 1 Then
                dco_popserver.Addentry dco_popserver.Getcurrententry
                dco_popusername.Addentry dco_popusername.Getcurrententry
                Me.Caption = "Frankmailer - Recieving Mail"
                checkmail socket, dco_popserver.Getcurrententry, 110, dco_popusername.Getcurrententry, txt_poppassword, timeout, False, App.Path
                fil_msgfiles.Refresh
                Unload frm_progress
                frm_main.Caption = "Frankmailer"
            End If
        Else
            MsgBox "Please wait...", vbOKOnly, "Frankmailer Busy"
        End If
    End Sub

'**********************************
'TIMEOUT TRIGGER
'**********************************
    Private Sub tim_timeout_Timer()
        timout_elapsed
    End Sub

'**********************************
'MESSAGE WINDOW EVENTS
'**********************************
    Private Sub fil_msgfiles_DblClick()
        If verify_file(App.Path & "\" & fil_msgfiles.filename) Then
            ShellExecute Me.hwnd, vbNullString, fil_msgfiles.filename, vbNullString, App.Path & "\", SW_SHOWNORMAL
        Else
            MsgBox "File " & fil_msgfiles.filename & " doesnt exist!", vbOKOnly, "Open error"
            fil_msgfiles.Refresh
        End If
    End Sub
    
'**********************************************
'VERIFY HIGHSCORE FILE EXISTS
'**********************************************
    Private Function verify_file(i_filename As String) As Boolean
        Dim fs
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.fileexists(i_filename) Then
            verify_file = True
        Else
            verify_file = False
        End If
    End Function

