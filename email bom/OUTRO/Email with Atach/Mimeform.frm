VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mail with Attachments!"
   ClientHeight    =   6825
   ClientLeft      =   1650
   ClientTop       =   2205
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   4755
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Frame frm1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   4755
      TabIndex        =   16
      Top             =   1080
      Width           =   3300
   End
   Begin VB.CommandButton delattach 
      Caption         =   "Del Attachment"
      Height          =   375
      Left            =   6455
      TabIndex        =   7
      Top             =   628
      Width           =   1600
   End
   Begin VB.ListBox AttachmentList 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   4755
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   120
      Width           =   3300
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   3855
   End
   Begin VB.CommandButton SendMimeConnect 
      Appearance      =   0  'Flat
      Caption         =   "Send"
      Height          =   375
      Left            =   120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   3975
   End
   Begin VB.ComboBox MailServer 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Text            =   "mailsend"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Attachment 
      BackColor       =   &H00000000&
      Caption         =   "Add Attachment"
      Height          =   375
      Left            =   4755
      TabIndex        =   6
      Top             =   628
      Width           =   1600
   End
   Begin VB.TextBox Tobox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "user@server.com"
      Top             =   594
      Width           =   2175
   End
   Begin VB.ComboBox Frombox 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   720
      TabIndex        =   3
      Text            =   "me@host.com"
      Top             =   1065
      Width           =   2175
   End
   Begin VB.TextBox Subjekt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MaxLength       =   78
      TabIndex        =   4
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox DataArrival 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3960
      Width           =   7935
   End
   Begin VB.TextBox Mailtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1965
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1920
      Width           =   7935
   End
   Begin VB.Label lblpcent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H00008080&
      Height          =   192
      Left            =   4332
      TabIndex        =   18
      Top             =   1140
      Width           =   228
   End
   Begin VB.Label Process 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   7935
   End
   Begin VB.Label ggg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      ForeColor       =   &H00008080&
      Height          =   192
      Left            =   108
      TabIndex        =   13
      Top             =   168
      Width           =   528
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H00008080&
      Height          =   192
      Left            =   396
      TabIndex        =   12
      Top             =   612
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H00008080&
      Height          =   192
      Left            =   228
      TabIndex        =   11
      Top             =   1080
      Width           =   408
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      ForeColor       =   &H00008080&
      Height          =   192
      Left            =   60
      TabIndex        =   10
      Top             =   1572
      Width           =   576
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fixes and additions by Luis Cantero from L.C. Enterprises:
'Attachments fixed
'Bad last character fixed
'Multiple attachment support added
'Progress bar added
'OLE Drag & Drop of files to Attachment list added
'Multiple recipients separated by comma support added
'Detect Internet connection function added
'Comments: lc-enterprises@usa.net - http://lcenterprises.net
'Original code by: Sebastian
'Thanks to AndrComm for the Base64 routines

                  
Dim bTrans As Boolean
Dim m_iStage As Integer
Dim Sock As Integer
Dim RC As Integer
Dim Bytes As Integer
Dim ResponseCode As Integer
Dim path As Variant

Dim objBase64 As New Base64

'*****************************************
'For the Mime File Field!
'*****************************************

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const OFN_READONLY = &H1
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_HIDEREADONLY = &H4
Const OFN_NOCHANGEDIR = &H8
Const OFN_SHOWHELP = &H10
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_NOVALIDATE = &H100
Const OFN_ALLOWMULTISELECT = &H200
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_PATHMUSTEXIST = &H800
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_CREATEPROMPT = &H2000
Const OFN_SHAREAWARE = &H4000
Const OFN_NOREADONLYRETURN = &H8000
Const OFN_NOTESTFILECREATE = &H10000
Const OFN_NONETWORKBUTTON = &H20000
Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Const OFN_EXPLORER = &H80000 ' new look commdlg
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Const OFN_SHAREFALLTHROUGH = 2
Const OFN_SHARENOWARN = 1
Const OFN_SHAREWARN = 0

Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'This is for the WaitforResponse Routine
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Dec's for the X disabling

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&

'For MIME processing
Dim Mime As Boolean
'For Filehandling
Dim arrRecipients As Variant
Dim CurrentE As Integer

Sub DisableX(frm As Form)

  Dim hMenu As Long
  Dim nCount As Long

    hMenu = GetSystemMenu(frm.hWnd, 0)
    nCount = GetMenuItemCount(hMenu)

    'Get rid of the Close menu and its separator
    Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
    Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)

    'Make sure the screen updates
    'our change
    DrawMenuBar frm.hWnd

End Sub

'***************************************************************
'Thanks to Luis Cantero for this Routines
Sub Startrek(frm As Form)

    GotoVal = frm.Height / 2
    For Gointo = 1 To GotoVal
        DoEvents
        frm.Height = frm.Height - 100
        frm.Top = (Screen.Height - frm.Height) \ 2
        If frm.Height <= 500 Then Exit For
    Next Gointo
horiz:
    frm.Height = 30
    GotoVal = frm.Width / 2
    For Gointo = 1 To GotoVal
        DoEvents
        frm.Width = frm.Width - 100
        frm.Left = (Screen.Width - frm.Width) \ 2
        If frm.Width <= 2000 Then Exit For
    Next Gointo

End Sub

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String

  Dim ofn As OPENFILENAME
  Dim A As Long

    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter & "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_CREATEPROMPT
    A = GetSaveFileName(ofn)
    If (A) Then
        SaveDialog = Left$(Trim$(ofn.lpstrFile), Len(Trim$(ofn.lpstrFile)) - 1)
      Else
        SaveDialog = ""
    End If

End Function

'***************************************************************
Private Sub Attachment_Click()

    path = SaveDialog(Me, "*.*", "Attach File", App.path)
    If path = "" Then Exit Sub
    AttachmentList.AddItem path
    Mime = True
    AttachmentList.ListIndex = AttachmentList.ListCount - 1

End Sub

Private Sub AttachmentList_Click()

    fSize = Int((FileLen(AttachmentList) / 1024) * 100 + 0.5) / 100
    AttachmentList.ToolTipText = AttachmentList & " (" & fSize & " KB)"

End Sub

Private Sub AttachmentList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    For i = 1 To Data.Files.Count
        If (GetAttr(Data.Files.Item(i)) And vbDirectory) = 0 Then AttachmentList.AddItem Data.Files.Item(i): Mime = True: AttachmentList.ListIndex = AttachmentList.ListCount - 1
    Next

End Sub

Private Sub delattach_Click()

    If AttachmentList.ListCount = 0 Or AttachmentList.ListIndex = -1 Then Exit Sub

    tmpIndex = AttachmentList.ListIndex
    AttachmentList.RemoveItem (AttachmentList.ListIndex)

    If AttachmentList.ListCount = 0 Then Mime = False: Attachment.ToolTipText = "Drag & Drop your attachments here" Else If AttachmentList.ListIndex = 0 Then AttachmentList.ListIndex = tmpIndex Else AttachmentList.ListIndex = tmpIndex - 1

End Sub

'***************************************************************
'Routine for connecting to the server
'***************************************************************
Private Sub SendMimeConnect_Click()

  ' Little Error check

    If Tobox = "" Or InStr(Tobox, "@") = 0 Then
        MsgBox "To: Is not correct!"
        Exit Sub
    End If

  Dim StartupData As WSADataType
  Dim SocketBuffer As sockaddr
  Dim IpAddr As Long
    
    'Ini the Winsocket
    RC = WSAStartup(&H101, StartupData)
    RC = WSAStartup(&H101, StartupData)
    
    'Open a free Socket (with this source code you can also
    'open several connections! Very useful for E-Mail Applications...)
    Sock = socket(AF_INET, SOCK_STREAM, 0)
    If Sock = SOCKET_ERROR Then
        Process = "Cannot Create Socket."
        Exit Sub
    End If

    If IsConnected2Internet = False Then
        Ret = MsgBox("No Internet connection has been detected, try to send Email anyway?", vbYesNo)
        If Ret = vbNo Then Exit Sub
    End If

    'Checks if the Hostname exists
    If RC = SOCKET_ERROR Then Exit Sub
    IpAddr = GetHostByNameAlias(MailServer)
    If IpAddr = -1 Then
        Process = "Unknown Host: " & MailServer
        Exit Sub
    End If

    'This part is responsible for the connection
    SocketBuffer.sin_family = AF_INET
    SocketBuffer.sin_port = htons(25)
    SocketBuffer.sin_addr = IpAddr
    SocketBuffer.sin_zero = String$(8, 0)
    
    RC = connect(Sock, SocketBuffer, Len(SocketBuffer))

    'If an error occured close the connection and
    'send an error message to the text window
    If RC = SOCKET_ERROR Then
        Process = "Cannot Connect to " & MailServer & GetWSAErrorString(WSAGetLastError())
        closesocket Sock
        Call WSACleanup
        Exit Sub
      Else
        Process = "Connected to " & MailServer
    End If

    'Select Receive Window
    RC = WSAAsyncSelect(Sock, DataArrival.hWnd, _
         ByVal &H202, ByVal FD_READ Or FD_CLOSE)
    If RC = SOCKET_ERROR Then
        Process = "Cannot Process Asynchronously."
        closesocket Sock
        Call WSACleanup
        Exit Sub
    End If

    bTrans = True
    m_iStage = 0
    DataArrival = ""

    ResponseCode = 220
    Call WaitForResponse

End Sub

Private Sub Exit_Click()

    On Error Resume Next
      Call Startrek(Me)

      closesocket Sock
      Call WSACleanup
      End

End Sub

Private Sub Form_Load()

    Call DisableX(Me)

End Sub

'***************************************************************
'Routine for arraving Data
'***************************************************************
Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim MsgBuffer As String * 2048

    On Error Resume Next
    
      If Sock > 0 Then
          'Receive up to 2048 chars
          Bytes = recv(Sock, ByVal MsgBuffer, 2048, 0)
        
          If Bytes > 0 Then
            
              DataArrival = DataArrival & _
                            MsgBuffer & _
                            vbCrLf
              'Scrolls down the Textbox
              DataArrival.SelStart = Len(DataArrival)
         
              If bTrans Then
                  'Checks if the Response code is correct
                  If ResponseCode = Left$(MsgBuffer, 3) Then
                      MsgBuffer = vbNullString
                      m_iStage = m_iStage + 1
                      Transmit m_iStage
                    Else
                      'If the Response Code is not right reset the connection
                      closesocket (Sock)
                      Call WSACleanup
                      Sock = 0
                      Process = "The Server responds with an unexpected Response Code!"
                      Exit Sub
                  End If
              End If

            ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
              closesocket (Sock)
              Call WSACleanup
              Sock = 0
          End If
      End If
      Refresh

End Sub

'***************************************************************
'Sends the E-Mail
'***************************************************************
Private Sub Transmit(iStage As Integer)

  Dim Helo As String
  Dim pos As Integer

    Select Case m_iStage

      Case 1:
        Helo = Frombox
        pos = Len(Helo) - InStr(Helo, "@")
        Helo = Right$(Helo, pos)

        ResponseCode = 250
        WinsockSendData ("HELO " & Helo & vbCrLf)

        Call WaitForResponse

      Case 2:
        ResponseCode = 250
        WinsockSendData ("MAIL FROM: <" & Trim$(Frombox) & ">" & vbCrLf)

        Call WaitForResponse

      Case 3:
        ResponseCode = 250
        WinsockSendData ("RCPT TO: <" & Trim$(arrRecipients(CurrentE)) & ">" & vbCrLf)

        Call WaitForResponse

      Case 4:
        ResponseCode = 354
        WinsockSendData ("DATA" & vbCrLf)

        Call WaitForResponse

      Case 5:
        ' Calls the routine to send the Header
        ResponseCode = 250
        Call SendMimetxt(Frombox, Trim$(arrRecipients(CurrentE)), Subjekt, Mailtxt)

        Call WaitForResponse

        'Finish the E-Mail sending process
      Case 6:
        ResponseCode = 221
        WinsockSendData ("QUIT" & vbCrLf)
        Call WaitForResponse

        Process = "Email has been sent!"
        frm2.Width = 3300
        lblpcent = "100%"

        DataArrival = ""

        m_iStage = 0
        If arrRecipients(CurrentE + 1) <> "" Then
            CurrentE = CurrentE + 1
            SendMimeConnect_Click
          Else
            bTrans = False
            CurrentE = 0
        End If
    End Select

End Sub

'***************************************************************
'Routine for sending a MIME txt
'***************************************************************
Sub SendMimetxt(txtFrom, txtTo, txtSubjekt, txtMail)

  Dim temp As Variant

    If Mime Then
        'Prepare the MIME Mail Header

        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'If you want additional Headers like Date,Message-Id,...etc. !
        'simply add them below                                      !
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        temp = "From: " & txtFrom & vbCrLf
        temp = temp & "To: " & txtTo & vbCrLf
        temp = temp & "Subject: " & txtSubjekt & vbCrLf

        'Do not change this Headers

        temp = temp & "Mime-Version: 1.0" & vbCrLf
        temp = temp & "Content-Type: multipart/mixed; boundary=" & Chr$(34) & "NextMimePart" & Chr$(34) & vbCrLf
        temp = temp & "Content-Transfer-Encoding: 7bit" & vbCrLf
        temp = temp & "This is a multi-part message in MIME format." & vbCrLf
        temp = temp & "--NextMimePart" & vbCrLf & "--NextMimePart" & vbCrLf & vbCrLf

        'Header plus Message
        temp = temp & Mailtxt

        'The following routine is necessary in order to be able to send
        'a dot on a single line without confusing the server
        '(Very Important, otherwhise the email might get truncated)

        tDot = 1
        For i = 1 To Len(temp)
            tDot = InStr(tDot + 4, temp, vbCrLf & "." & vbCrLf)
            If tDot = 0 Then Exit For
            temp = Mid$(temp, 1, tDot + 2) & Chr$(0) & Mid$(temp, tDot + 3)
            DoEvents
        Next

        'Send the Mime Header and the Message
        'send message body in steps in case it is too large
        For i = 1 To Len(temp) Step 8192
            ss = Trim$(Mid$(temp, i, 8192))
            WinsockSendData (ss)

            frm2.Width = (2400 / 100) * ((i + Len(ss)) * 100 / Len(temp))
            lblpcent = Int((i + Len(ss)) * 100 / Len(temp)) & "%"
            If Cancelflag Then Exit For
            Process = "Sending message body... " & i + Len(ss) & " Bytes from " & Len(temp)
            DoEvents
        Next

        'Call Attachment Routine
        SendMimeAttachment

      Else
        'Send the E-Mail without Attachment

        temp = "From: " & txtFrom & vbCrLf
        temp = temp & "To: " & txtTo & vbCrLf
        temp = temp & "Subject: " & txtSubjekt & vbCrLf
        temp = temp & txtMail

        'The following routine is necessary in order to be able to send
        'a dot on a single line without confusing the server
        '(Very Important, otherwhise the email might get truncated)

        tDot = 1
        For i = 1 To Len(temp)
            tDot = InStr(tDot + 4, temp, vbCrLf & "." & vbCrLf)
            If tDot = 0 Then Exit For
            temp = Mid$(temp, 1, tDot + 2) & Chr$(0) & Mid$(temp, tDot + 3)
            DoEvents
        Next

        'send message body in steps in case it is too large
        For i = 1 To Len(temp) Step 8192
            ss = Trim$(Mid$(temp, i, 8192))
            WinsockSendData (ss)

            frm2.Width = (2400 / 100) * ((i + Len(ss)) * 100 / Len(temp))
            lblpcent = Int((i + Len(ss)) * 100 / Len(temp)) & "%"
            If Cancelflag Then Exit For
            Process = "Sending message body... " & i + Len(ss) & " Bytes from " & Len(temp)
            DoEvents
        Next

        'Send Data and finish it!
        WinsockSendData (vbCrLf & "." & vbCrLf)
    End If

End Sub

'**************************************************************
'NEW! Waits until time out, while waiting for response
'**************************************************************
Private Sub WaitForResponse()

  Dim Start As Long
  Dim Tmr As Long

    'Works with an Api Declaration because it's more presice

    Start = timeGetTime
    While Bytes > 0
        Tmr = timeGetTime - Start
        
        DoEvents ' Let System keep checking for incoming response
     
        'Wait 20 (20000 Miliseconds) seconds for response
        If Tmr > 20000 Then
            Process = "SMTP service error, timed out while waiting for response"
            'End
        End If
    Wend

End Sub

'***************************************************************
'Routine for sending a MIME Attachment
'***************************************************************
Private Sub SendMimeAttachment()

  Dim FileIn As Long
  Dim temp As Variant
  Dim s As Variant

  Dim TempArray() As Byte
  Dim Encoded() As Byte
  Dim strFile As String
  Dim strFile1 As String * 32768

    For Iat = 0 To AttachmentList.ListCount - 1
        path = AttachmentList.List(Iat)

        Mimefilename = Trim$(Right$(path, Len(path) - InStrRev(path, "\")))

        'Gets the next free filenumber
        FileIn = FreeFile

        'Preparing the Mime Header
        temp = vbCrLf & "--NextMimePart" & vbNewLine
        temp = temp & "Content-Type: application/octet-stream; name=" & Chr$(34) & Mimefilename & Chr$(34) & vbNewLine
        temp = temp & "Content-Transfer-Encoding: base64" & vbNewLine
        temp = temp & "Content-Disposition: attachment; filename=" & Chr$(34) & Mimefilename & Chr$(34) & vbNewLine

        WinsockSendData (temp & vbCrLf)

        'Open Base64 Input File
        Open path For Binary Access Read As FileIn
        If GetSetting(App.Title, "Settings", "Too big", "") <> "True" Then
            If LOF(FileIn) > 2097152 Then
                fSize = Int((LOF(FileIn) / 1048576) * 100 + 0.5) / 100
                Setu = MsgBox("The current file is " & fSize & " MB of size, extracting from it could take a few minutes, Click Yes to go ahead, No to skip it or Cancel if you don't want to get this message again", vbYesNoCancel)
                If Setu = vbYes Then GoTo Cont
                If Setu = vbNo Then Close (FileIn): GoTo Anoth Else SaveSetting App.Title, "Settings", "Too big", "True"
            End If
        End If

Cont:

        frm2.Visible = True
        Process = "Loading """ & AttachmentList.List(Iat) & """"
        Do While Not EOF(FileIn)
            If LOF(FileIn) = 0 Then GoTo Anoth
            Get FileIn, , strFile1
            strFile = strFile & Mid$(strFile1, 1, Len(strFile1) - (Loc(FileIn) - LOF(FileIn)))
            strFile1 = ""
            DoEvents

            frm2.Width = (3300 / 100) * (Len(strFile) * 50 / LOF(FileIn))
            lblpcent = Int(Len(strFile) * 50 / LOF(FileIn)) & "%"

            If Cancelflag Then Close FileIn: Exit Sub
        Loop
        Close FileIn

        If strFile = "" Then Exit Sub

        objBase64.Str2ByteArray strFile, TempArray
        objBase64.EncodeB64 TempArray, Encoded
        objBase64.Span 76, Encoded, TempArray

        strFile = ""

        s = StrConv(TempArray, vbUnicode)

        For i = 1 To Len(s) Step 8192
            ss = Trim$(Mid$(s, i, 8192))
            
            tmpServerSpeed = 150 'milliseconds
            Start = timeGetTime
            Do
                DoEvents
            Loop Until timeGetTime >= Start + tmpServerSpeed * 20
                
            WinsockSendData (ss)

            frm2.Width = 1650 + (3300 / 100) * ((i + Len(ss)) * 50 / Len(s))
            lblpcent = 50 + Int((i + Len(ss)) * 50 / Len(s)) & "%"

            Process = "Sending " & Mimefilename & "... " & i + Len(ss) & " Bytes from " & Len(s)
            DoEvents
        Next

        'Send the last part of the MIME Body
Anoth:
        s = ""
    Next
    WinsockSendData (vbCrLf & "--NextMimePart--" & vbCrLf)
    WinsockSendData (vbCrLf & "." & vbCrLf)

End Sub

Private Sub WinsockSendData(DatatoSend As String)

  Dim RC As Integer
  Dim MsgBuffer As String * 8192

    MsgBuffer = DatatoSend

    'You can open more than one connection!
    RC = send(Sock, ByVal MsgBuffer, Len(DatatoSend), 0)
    
    'If an error occurs send an error message and
    'reset the winsock
    If RC = SOCKET_ERROR Then
        Process = "Cannot Send Request." & Str$(WSAGetLastError()) & _
                  GetWSAErrorString(WSAGetLastError())
        closesocket Sock
        Call WSACleanup
        Exit Sub
    End If

End Sub

Private Sub Tobox_Change()

    arrRecipients = Split(Tobox, ",")

End Sub

Function IsConnected2Internet() As Boolean

    On Error Resume Next
      'IsConnected = InternetGetConnectedState(0&, 0&) 'Doesn't work with older versions of Wininit.dll

      If MyIP = "127.0.0.1" Or MyIP = "" Then IsConnected2Internet = False Else IsConnected2Internet = True

End Function

':) Ulli's Code Formatter V2.0 (22.04.2001 16:57:05) 95 + 574 = 669 Lines
