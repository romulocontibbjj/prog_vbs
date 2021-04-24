Attribute VB_Name = "mod_email"
Option Explicit

'************************************
'PRIVATE CONSTANTS
'************************************
    Private Const timout_length = 10000

'************************************
'PRIVATE TYPE DECLARATIONS
'************************************
    Private Type mailmsg
        rxdate As String
        rxtime As String
        message As String
    End Type
    
'************************************
'PRIVATE VARIABLES
'************************************
    Private task_timeout As Boolean

'************************************
'PUBLIC VARIABLES
'************************************
    Public noofmsgs As Integer
    Public mailmsgs() As mailmsg
    Public mailbusy As Boolean

'************************************
'PUBLIC FUNCTIONS
'************************************
    'SENDS EMAIL
    Public Function sendmail(ByVal socket As Winsock, mailsmtp As String, mailport As Long, mailto As String, _
                            mailfrom As String, mailsubject As String, mailbody As String, ByVal timeout As Timer _
                            ) As Boolean
        Load frm_progress
        frm_progress.Caption = "Sending mail"
        frm_progress.pro_progress.Max = 5
        frm_progress.pro_progress = 0
        frm_progress.Show
        mailbusy = True
        If socket.State <> sckClosed Then socket.Close
        socket.Connect mailsmtp, mailport
        
        reset_timeout timeout
        start_timeout timeout
        
        While socket.State <> sckConnected
            DoEvents
            frm_progress.lbl_step = "Connecting to " & mailsmtp
            If task_timeout Then GoTo sendmailtimeout
        Wend

        If findmessage(socket, "220", timeout) Then
            frm_progress.pro_progress = 1
            frm_progress.lbl_step = "Introducing"
            socket.SendData "helo " & mailsmtp & vbCrLf
            If findmessage(socket, "250", timeout) Then
                frm_progress.pro_progress = 2
                frm_progress.lbl_step = "Sending mail"
                socket.SendData "mail from:" & mailfrom & vbCrLf
                If findmessage(socket, "250", timeout) Then
                    frm_progress.pro_progress = 3
                    socket.SendData "rcpt to:" & mailto & vbCrLf
                    If findmessage(socket, "250", timeout) Then
                        frm_progress.pro_progress = 4
                        socket.SendData "DATA" & vbCrLf
                        If findmessage(socket, "354", timeout) Then
                            socket.SendData "From: " & mailfrom & vbCrLf
                            socket.SendData "Subject: " & mailsubject & vbCrLf
                            socket.SendData "To: " & mailto & vbCrLf & vbCrLf
                            socket.SendData mailbody & vbCrLf
                            socket.SendData vbCrLf & "." & vbCrLf
                            If findmessage(socket, "25", timeout) Then
                                frm_progress.pro_progress = 5
                                frm_progress.lbl_step = "Mail sent"
                                If socket.State = sckConnected Then socket.Close
                                reset_timeout timeout
                                sendmail = True
                                mailbusy = False
                                Unload frm_progress
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
sendmailtimeout:
        mailbusy = False
        If socket.State = sckConnected Then socket.Close
    End Function
    
    'CHECKS EMAIL
    Public Function checkmail(ByVal socket As Winsock, mailpop3 As String, mailport As Long, mailuser As String, _
                            mailpass As String, ByVal timeout As Timer, delfromserver As Boolean, _
                            mailpath As String) As Integer
        Load frm_progress
        frm_progress.Caption = "Recieving Mail"
        frm_progress.pro_progress.Max = 3
        frm_progress.pro_progress = 0
        frm_progress.Show
        mailbusy = True
        Dim msgnumber As Integer
        Dim currentmsg As Integer
        Dim gotnoofmsgs As Boolean
                            
        If socket.State <> sckClosed Then socket.Close
        socket.Connect mailpop3, mailport
        
        reset_timeout timeout
        start_timeout timeout
        
        While socket.State <> sckConnected
            DoEvents
            frm_progress.lbl_step = "Connecting to " & mailpop3
            If task_timeout Then GoTo checkmailtimeout
        Wend
        
        If findmessage(socket, "+OK", timeout) Then
            frm_progress.pro_progress = 1
            frm_progress.lbl_step = "Authenticating user"
            socket.SendData "user " & mailuser & vbCrLf
            If findmessage(socket, "+OK", timeout) Then
                frm_progress.pro_progress = 2
                frm_progress.lbl_step = "Authenticating password"
                socket.SendData "pass " & mailpass & vbCrLf
                If findmessage(socket, "+OK", timeout) Then
                    frm_progress.pro_progress = 3
                    frm_progress.lbl_step = "Checking for messages"
                    socket.SendData "list 1" & vbCrLf
                    If findmessage(socket, "+OK", timeout) Then
                        msgnumber = msgnumber + 1
                        While Not gotnoofmsgs
                            DoEvents
                            socket.SendData "list " & Str(msgnumber) & vbCrLf
                            If findmessage(socket, "+OK", timeout) Then
                                frm_progress.lbl_step = "Found " & Str(msgnumber) & " message(s)..."
                                msgnumber = msgnumber + 1
                            Else
                                frm_progress.lbl_step = "Found " & Str(msgnumber - 1) & " message(s)"
                                msgnumber = msgnumber - 1
                                gotnoofmsgs = True
                            End If
                        Wend
                        If msgnumber > 0 Then
                            frm_progress.lbl_step = "Downloading messages"
                            frm_progress.pro_progress.Max = msgnumber
                            checkmail = msgnumber
                            ReDim mailmsgs(msgnumber) As mailmsg
                            For currentmsg = 1 To msgnumber
                                DoEvents
                                frm_progress.lbl_step = "Downloading message " & currentmsg
                                frm_progress.pro_progress = currentmsg
                                socket.SendData "retr " & Str(currentmsg) & vbCrLf
                                mailmsgs(currentmsg).message = getmailmsg(socket)
                                mailmsgs(currentmsg).rxdate = removebad(Date)
                                mailmsgs(currentmsg).rxtime = removebad(Time)
                                frm_progress.lbl_step = "Deleting message " & currentmsg & " from server"
                                socket.SendData "DELE " & Str(currentmsg) & vbCrLf
                                If findmessage(socket, "+OK", timeout) Then
                                    frm_progress.lbl_step = "Saving message " & currentmsg & " to message folder"
                                    Open mailpath & "\" & mailmsgs(currentmsg).rxdate & "@" & mailmsgs(currentmsg).rxtime & Str(currentmsg) & ".eml" For Output As #1
                                        Print #1, mailmsgs(currentmsg).message
                                    Close #1
                                End If
                            Next currentmsg
                        End If
                        socket.SendData "QUIT" & vbCrLf
                        If findmessage(socket, "+OK", timeout) Then
                            If socket.State = sckConnected Then socket.Close
                        End If
                        mailbusy = False
                        Exit Function
                    End If
                    checkmail = 0
                    mailbusy = False
                    Exit Function
                End If
            End If
        End If

checkmailtimeout:
        checkmail = -1
        mailbusy = False
        If socket.State = sckConnected Then socket.Close
    End Function
    
    'TIMES OUT CURRENT TASK
    Public Sub timout_elapsed()
        task_timeout = True
    End Sub

'************************************
'BUFFERS INCOMING SOCKET DATA
'************************************
    'MESSAGE HANDLER
    Private Function findmessage(ByVal socket As Winsock, findstring As String, timeout As Timer) As Boolean
        Dim rxbuffer As String
        Dim strdata As String
        
        socket.GetData rxbuffer, vbString
        reset_timeout timeout
        start_timeout timeout
        
        While InStr(1, rxbuffer, findstring, vbTextCompare) = 0
            DoEvents
            If socket.State = sckConnected Then
                socket.GetData strdata, vbString
                rxbuffer = rxbuffer & strdata
            Else
                GoTo findmessagetimeout
            End If
            If task_timeout Then GoTo findmessagetimeout
        Wend
        findmessage = True
        Exit Function
findmessagetimeout:
    End Function

    'MAIL HANDLER
    Private Function getmailmsg(ByVal socket As Winsock) As String
        Dim rxbuffer As String
        Dim strdata As String
        
        socket.GetData rxbuffer, vbString
        
        While InStr(1, rxbuffer, vbCrLf & "." & vbCrLf, vbTextCompare) = 0
            DoEvents
            If socket.State = sckConnected Then
                socket.GetData strdata, vbString
                rxbuffer = rxbuffer & strdata
            End If
        Wend
        getmailmsg = rxbuffer
        Exit Function
findmessagetimeout:
    End Function

'************************************
'PRIVATE FUNCTIONS
'************************************
    'RESETS A TIMER
    Private Sub reset_timeout(timeout As Timer)
        task_timeout = False
        timeout.enabled = False
        timeout.Interval = timout_length
    End Sub

    'STARTS A TIMER
    Private Sub start_timeout(timeout As Timer)
        timeout.Interval = timout_length
        timeout.enabled = True
    End Sub

    'PRIVATE SUB REMOVE ILLEGAL CHARACTERS
    Private Function removebad(strdata As String) As String
        strdata = Replace(strdata, "\", "-")
        strdata = Replace(strdata, "/", "-")
        strdata = Replace(strdata, ":", "-")
        strdata = Replace(strdata, "*", "-")
        strdata = Replace(strdata, "?", "-")
        strdata = Replace(strdata, """", "-")
        strdata = Replace(strdata, "<", "-")
        strdata = Replace(strdata, ">", "-")
        strdata = Replace(strdata, "|", "-")
        removebad = strdata
    End Function
