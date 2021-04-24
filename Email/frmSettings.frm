VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3900
      Left            =   45
      TabIndex        =   4
      Top             =   15
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   6879
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sender E-Mail"
      TabPicture(0)   =   "frmSettings.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOK"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClose"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "SMTP Server"
      TabPicture(1)   =   "frmSettings.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   350
         Left            =   -70470
         TabIndex        =   3
         Top             =   2700
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   350
         Left            =   -71820
         TabIndex        =   2
         Top             =   2700
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         Height          =   1320
         Left            =   -74790
         TabIndex        =   17
         Top             =   1305
         Width           =   5550
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1545
            TabIndex        =   1
            Top             =   765
            Width           =   3915
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1545
            TabIndex        =   0
            Top             =   255
            Width           =   3915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Address"
            Height          =   195
            Left            =   165
            TabIndex        =   19
            Top             =   765
            Width           =   1260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   315
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "SMTP Server"
         Height          =   3330
         Left            =   120
         TabIndex        =   10
         Top             =   405
         Width           =   5655
         Begin VB.TextBox txtSMTPServer 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1575
            TabIndex        =   5
            Top             =   840
            Width           =   3885
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            Height          =   325
            Left            =   3135
            TabIndex        =   8
            Top             =   2145
            Width           =   1065
         End
         Begin VB.CheckBox chkProxy 
            Caption         =   "Proxy"
            Height          =   225
            Left            =   195
            TabIndex        =   6
            Top             =   1755
            Width           =   1020
         End
         Begin VB.TextBox txtProxy 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1575
            TabIndex        =   7
            Top             =   1710
            Width           =   3885
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Close"
            Height          =   325
            Left            =   4380
            TabIndex        =   9
            Top             =   2145
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please Specify the SMTP Server for outgoing e-mail"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   360
            Width           =   4440
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Server"
            Height          =   195
            Left            =   225
            TabIndex        =   15
            Top             =   870
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(IP Addres (or) Server Name)"
            Height          =   195
            Left            =   1680
            TabIndex        =   14
            Top             =   1215
            Width           =   2580
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Ex. mx2.hotmail.com)"
            Height          =   195
            Left            =   1845
            TabIndex        =   13
            Top             =   1470
            Width           =   1965
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   " if you are using Proxy server, you must specify the Proxy Server name (or) IP Address"
            Height          =   465
            Left            =   795
            TabIndex        =   12
            Top             =   2595
            Width           =   4650
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   2625
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkProxy_Click()
Dim ret As String
If chkProxy.Value = 1 Then
    txtProxy.Enabled = True
    txtSMTPServer.Enabled = False
    ret = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer")
    If ret <> "" Then
        txtProxy.Text = (Mid(ret, 1, InStr(ret, ":") - 1))
    End If
Else
    txtProxy.Enabled = False
    txtSMTPServer.Enabled = True
End If
End Sub

Private Sub cmdApply_Click()
Dim ret As Boolean
If chkProxy.Value = 1 And Trim(txtProxy.Text) <> "" Then
    ret = SMTPKEY(txtProxy.Text, 3)
ElseIf Trim(txtSMTPServer.Text) <> "" Then
    ret = SMTPKEY(txtSMTPServer.Text, 2)
Else
   ret = SMTPKEY("FIND", 1)
End If

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim ret As Boolean
If Trim(txtName.Text) = "" Then
        MsgBox "Please enter your name", vbInformation + vbOKOnly, "Smart Easy E-Mail"
        txtName.SetFocus
        Exit Sub
End If
If Trim(txtEmail.Text) <> "" Then
    If EmailValid(txtEmail.Text) = False Then
        MsgBox "Not a valid email address", vbInformation + vbOKOnly, "Smart Easy E-Mail"
        txtEmail.SetFocus
        Exit Sub
    End If
Else
     MsgBox "Please enter your E-Mail address", vbInformation + vbOKOnly, "Smart Easy E-Mail"
     txtEmail.SetFocus
End If

If Trim(txtName.Text) <> "" And Trim(txtEmail.Text) <> "" Then
   ret = SenderEmail(txtName.Text, txtEmail.Text)
   frmMain.txtFrom.Text = txtName.Text
   frmMain.txtFrom.Tag = txtEmail.Text
End If
End Sub

Private Sub Form_Load()
Dim ret As String
'--------SMTP Server
   ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "SMTPIP")
    If ret <> "" Then
        If Split(ret, "%")(1) = "2" Then
            txtSMTPServer.Text = Split(ret, "%")(0)
        ElseIf Split(ret, "%")(1) = "3" Then
            txtProxy.Text = Split(ret, "%")(0)
            chkProxy.Value = 1
        Else
            txtSMTPServer.Text = ""
            txtProxy.Text = ""
        End If
    End If
'--------Sender Info
   ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "SENDER")
    If ret <> "" Then
        txtName.Text = Split(ret, "%")(0)
        txtEmail.Text = Split(ret, "%")(1)
    End If
    SSTab1.Tab = 0
End Sub
