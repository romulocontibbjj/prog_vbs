VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP - Customizado"
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14025
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstFtpFiles 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      ItemData        =   "Form1.frx":030A
      Left            =   9540
      List            =   "Form1.frx":030C
      TabIndex        =   21
      Top             =   3240
      Width           =   4470
   End
   Begin VB.FileListBox FlsLst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   6345
      TabIndex        =   20
      Top             =   3330
      Width           =   3075
   End
   Begin VB.DirListBox DirLst 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   6345
      TabIndex        =   19
      Top             =   2115
      Width           =   3075
   End
   Begin VB.DriveListBox DrvLst 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6345
      TabIndex        =   18
      Top             =   1755
      Width           =   3075
   End
   Begin MSComctlLib.TreeView trvFtpFldrs 
      Height          =   2445
      Left            =   9540
      TabIndex        =   17
      Top             =   405
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   4313
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7965
      Top             =   6705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030E
            Key             =   "ClosFldr"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0760
            Key             =   "OpenFldr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BB2
            Key             =   "sFile"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer4 
      Left            =   8550
      Top             =   6210
   End
   Begin VB.Timer TimWorkin 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10305
      Top             =   5130
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Lembrar-me neste computador"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2745
      TabIndex        =   6
      Top             =   1980
      Width           =   3480
   End
   Begin VB.TextBox TxtPW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2745
      TabIndex        =   5
      Top             =   1485
      Width           =   3480
   End
   Begin VB.TextBox TxtUsrNm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2745
      TabIndex        =   4
      Text            =   "anonymous"
      Top             =   855
      Width           =   3480
   End
   Begin VB.TextBox TxtFtp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2745
      TabIndex        =   3
      Text            =   "I.e. (ftp://ftp.microsoft.com)"
      Top             =   225
      Width           =   3480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   6660
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   6210
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   5760
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7965
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label LblCon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Conectar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   30
      TabIndex        =   7
      Top             =   2880
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape43 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   5355
      Width           =   2910
   End
   Begin VB.Label LbNewFldr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criar Nova Pasta"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   30
      TabIndex        =   22
      Top             =   5355
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape23 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   5850
      Width           =   2910
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   30
      TabIndex        =   16
      Top             =   5850
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos FTP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   10830
      TabIndex        =   15
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maquina Local"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6930
      TabIndex        =   14
      Top             =   1395
      Width           =   1860
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pastas FTP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   10605
      TabIndex        =   13
      Top             =   45
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   2880
      Width           =   2910
   End
   Begin VB.Label LblDnldFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Download Arquivo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   30
      TabIndex        =   10
      Top             =   3870
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   3870
      Width           =   2910
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   3375
      Width           =   2910
   End
   Begin VB.Label LblUploadFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Arquivo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   45
      TabIndex        =   8
      Top             =   3375
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblDelFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deletar Arquivo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   30
      TabIndex        =   12
      Top             =   4365
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   4365
      Width           =   2910
   End
   Begin VB.Label LblRnmFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Renomear Arquivo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   30
      TabIndex        =   11
      Top             =   4860
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   30
      Top             =   4860
      Width           =   2910
   End
   Begin VB.Label LblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status : Nao Conectado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6585
      TabIndex        =   9
      Top             =   7320
      Width           =   2385
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   465
      Left            =   1530
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   12480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      X1              =   1170
      X2              =   3150
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   1170
      X2              =   3150
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   1170
      X2              =   3150
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Senha::"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   450
      TabIndex        =   2
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Usuario :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   450
      TabIndex        =   1
      Top             =   855
      Width           =   1740
   End
   Begin VB.Label LblUsrNm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Dominio :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   450
      TabIndex        =   0
      Top             =   225
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   0
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu MnuFiles 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MnuFilesExit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMin 
         Caption         =   "Minimize"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This will be used to fire a URL in the default browser, to support us. "Thanx in advance"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim FSO As New Scripting.FileSystemObject
Dim Nfile, X, Y, Z As Integer
Public X1, X2, X3, X4, StrData, Status, FtpFile, LoclFilNm, NewName, _
NewFldrNm, DnldFileNm, UpldFileNm, UsrFriendly As String
Public vtData, Ifile As Variant ' Data variable.
Private Sub DisplayData()
On Error GoTo MyErr
'display the data .
'Use the split function to parse the data for us
'Place dirs in the treeview and files in the listview
Dim sFTPFiles() As String
    Select Case Status
        Case "DIR"
            'clear the listview
            LstFtpFiles.Clear
            DoEvents ' time to clear the listview
            Status = Trim("")
            If Len(StrData) > 0 Then
               sFTPFiles = Split(StrData, vbCrLf)
                Dim i As Integer
                For i = 0 To UBound(sFTPFiles) - 1
                    If Len(sFTPFiles(i)) > 0 Then 'check for 0 len string
                        If InStr(1, sFTPFiles(i), "/") Then 'it is a dir
                            'Put the dir under the selected node
                            'when we create a key we place an _ in front of the key
                            'This is because a key must start with a non numeric
                            Dim oNode As Node
                                Set oNode = trvFtpFldrs.Nodes.Add(trvFtpFldrs.SelectedItem.Key, _
                                tvwChild, _
                                "_" & Left(sFTPFiles(i), Len(sFTPFiles(i)) - 1), _
                                Left(sFTPFiles(i), Len(sFTPFiles(i)) - 1), _
                                "ClosFldr")
                            If Not (oNode Is Nothing) Then
                                oNode.EnsureVisible
                                oNode.Parent.Image = "OpenFldr"
                            End If
                        Else 'it is a file
                            LstFtpFiles.AddItem sFTPFiles(i)
                        End If
                    End If
                Next i
            End If
        Case "CD"
            'we changed dirs so list any files and subdirs located in the selected dir
            Status = "DIR"
            'Inet1.Execute , "PWD"
            Inet1.Execute , "DIR" 'if the dir is empty it could take a while for this to complete
        Case "REN"
            ' do a Dir so you can see that the file name did change
            Status = "DIR"
            Inet1.Execute , "DIR"
        Case "DEL"
            ' do a Dir so you can see that the file was deleted
            Status = "DIR"
            Inet1.Execute , "DIR"
        Case "GET"
            Status = Trim("")
            'update the file control to show the file
            FlsLst.Refresh
        Case "PUT"
            'we want to do a dir to show the file is now there
            Status = "DIR"
            Inet1.Execute , "DIR"
        Case "DEL"
            'we want to do a dir to show that the file is gone
            Status = "DIR"
            Inet1.Execute , "DIR"
        Case "REN"
            'we want to do (DIR) to show the new file name
            Status = "DIR"
            Inet1.Execute , "DIR"
        Case "MKDIR"
            'we want to do (dir) to show the new Dir
            Status = "DIR"
            Inet1.Execute , "DIR"
    End Select
Exit Sub
MyErr:
    If Err.Number = 35602 Then
        'we already have a node with that Key so we will resume next.
        'This will only cause a problem if you have dirs with the same name located at different levels
        'This is because we use the dir name for the key when we create the node in the treeview.
        Resume Next
    Else
        MsgBox Str(Err.Number) & ": " & Err.Description, vbOKOnly, "Error"
    End If
End Sub
Private Sub Check1_Click()
If TxtFtp.Text = Trim("") Then
    LblStatus = ("Status : Informe o ftp remoto ftp:// server")
    TxtFtp.SetFocus
    Exit Sub
    Else
    LblStatus = ("Status : ...............")
    End If
    If Check1.Value = 1 Then
Open App.Path & "/log.txt" For Output As #Nfile
    Print #Nfile, TxtFtp
    Print #Nfile, TxtUsrNm
    Print #Nfile, TxtPW
    Print #Nfile, Check1.Value
Close #Nfile
    Else
Kill App.Path & "/log.txt"
    End If
End Sub
Private Sub DirLst_Change()
On Error Resume Next
    'change the flslst control to show the correct dir
    FlsLst.Path = DirLst.Path
End Sub
Private Sub DrvLst_Change()
'Errors like : Drive unavailable
On Error Resume Next
    'change the dirlst control to the correct drive in DrvLst
    DirLst.Path = DrvLst.Drive
End Sub
Private Sub flslst_Click()
If Inet1.URL = "" Then Exit Sub
LoclFilNm = FlsLst.FileName
End Sub
Private Sub FlsLst_DblClick()
'If a local file was double-clicked then upload it
If Inet1.URL = Trim("") Then
        Beep
        LblStatus.Caption = ("Status : Não Conectado")
        Exit Sub
    End If
Call LblUploadFile_Click
End Sub
Private Sub Form_Load()
Shape1.Height = 100
Shape1.Width = 100
Shape1.Left = 1
Shape1.Top = 1
Me.BackColor = RGB(0, 0, 255)
Timer1.Enabled = 1
X = 0: Z = 0: Y = 0

    Set trvFtpFldrs.ImageList = ImageList1
    trvFtpFldrs.Nodes.Add , , "ROOT", "Root", "ClosFldr"
    trvFtpFldrs.Nodes("ROOT").Selected = True 'select the root node
    
    'Write Logs
Nfile = FreeFile
'If 'log.txt' doesn't exist then create it.
If FSO.FileExists(App.Path & "/log.txt") = True Then
Open App.Path & "/log.txt" For Input As #Nfile
'If 'log.txt' is empty then write to it.
    If FileLen(App.Path & "/log.txt") = 0 Then
    Close #Nfile
    Exit Sub
    Else
    Line Input #Nfile, X1
    Line Input #Nfile, X2
    Line Input #Nfile, X3
    Line Input #Nfile, X4
Close #Nfile
TxtFtp.Text = X1
TxtUsrNm.Text = X2
TxtPW.Text = X3
Check1.Value = X4
Exit Sub
End If
Else
Set Ifile = FSO.CreateTextFile(App.Path & "/log.txt", True)
End If
Ifile.Close
Close #Nfile
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Trigers the MnuFiles Object
If Button = 2 Then PopupMenu MnuFiles
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set vtData = Nothing
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
'Here is where the connections are being worked.
'Start the timWorking timer to show that we are in a call.
    TimWorkin.Enabled = True
    'Show that we are working.
    LblStatus.Caption = ("Trabalhando na última requisição")
    'check all the states and list them in the listbox
    Select Case State
        Case Is = 1
        LblStatus.Caption = "O controle está procurando o endereço IP do computador host especificado. (icHostResolvingHost)"
    Case Is = 2
        LblStatus.Caption = "O controle encontrou o endereço IP do computador host especificado. (icHostResolved)"
    Case Is = 3
        LblStatus.Caption = "O controle esta se conectando com o computador host. (icConnecting)"
    Case Is = 4
        LblStatus.Caption = "O controle se conectou ao computador host. (icConnected)"
    Case Is = 5
        LblStatus.Caption = "O controle esta enviando uma solicitação para o computador host. (icRequesting)"
    Case Is = 6
        LblStatus.Caption = "O controle enviou cum requisição com sucesso. (icRequestSent)"
    Case Is = 7
        LblStatus.Caption = "O controle está recebendo uma resposta do computador host. (icReceivingResponse)"
    Case Is = 8
        LblStatus.Caption = "O controle recebu com com êxito uma resposta do computador host. (icResponseReceived)"
    Case Is = 9
        LblStatus.Caption = "O controle esta desconectando do computador host. (icDisconnecting)"
    Case Is = 10
        LblStatus.Caption = "O controle desconectou com êxito do computador host. (icDisconnected)"
    Case Is = 11
        LblStatus.Caption = "Ocorreu um erro na comunicação com o computador host. (icError)"
        'Or you can try to caption the error.
        'LblStatus.Caption = "Response Data from last request: " & "Response code: " & Str(Inet1.ResponseCode) & _
        '    "   Response Info: " & Inet1.ResponseInfo
        'Inet1.Cancel
    Case Is = 12  'request completed, now get the data
    LblStatus.Caption = "Status : Você esta Conectado. (icResponseCompleted) - " & UsrFriendly
         StrData = ""
        Dim Bdone As Boolean
        Bdone = False
        ' Get first chunk.
        vtData = Inet1.GetChunk(1024, icString)
        Do While Not Bdone
           StrData = StrData & vtData
           DoEvents
           ' Get next chunk.
           vtData = Inet1.GetChunk(1024, icString)
           If Len(vtData) = 0 Then
              Bdone = True 'If done (vtData transfer's done) ,then it's not gonna loop anymore.
              LblStatus.ForeColor = vbBlue
           End If
        Loop
        'we use a timer because if we called the DisplayData method directly the
        'StateChanged event would still be on the stack. This would cause a problem when the
        'DisplayData method did an Inet1.Execute. We would get a 35764 "Still executing last request" error.
        'So the timer lets the StateChanged event to finish and be removed from the stack.
        'Also note that the application must not be busy while retrieving your data using (Timer4), so if you clicked
        'Connect (LblCon) while the timer1 , timer2 and timer3 still working then you will get the error 35764 "Still executing last request"
        'I've made this on purpose to demonstrate the case clearly, so you have 2 choices :
        '1) Find a way to trigger timer4 after the application's timers are being not busy.
        '2) Find a way to force the use not to use the form untill it loads completely.
        'In my example here, i'm forcing the user not to use the controls untill the forms load completely.
        Timer4.Interval = 100
        Timer4.Enabled = True
    End Select
End Sub
Private Sub Label6_Click()
If Inet1.URL = Trim("") Then End
    'close the FTP connection
    If Inet1.StillExecuting Then
        'try to cancel the request
        Inet1.Cancel 'It may take a while to cancel a request
    End If
    Status = "CLOSE"
    Inet1.Execute , "CLOSE"
    LstFtpFiles.Clear
    trvFtpFldrs.Nodes.Clear
    'we just cleared the nodes so we need to add the root node back
    'set the root node for the treeview
    trvFtpFldrs.Nodes.Add , , "ROOT", "Root", "ClosFldr"
    trvFtpFldrs.Nodes("ROOT").Selected = True 'select the root node
    End
End Sub
Private Sub Label8_Click()
'Open URL in the default browser.
 ShellExecute 0, vbNullString, "http://evry1falls.freevar.com/VB6/", vbNullString, vbNullString, vbNormalFocus
End Sub
Private Sub LblCon_Click()
  'make sure we don't have any old data in the listview or treeview
    LstFtpFiles.Clear
    trvFtpFldrs.Nodes.Clear
    'we just cleared the nodes so we need to add the root node back
    'set the root node for the treeview
    trvFtpFldrs.Nodes.Add , , "ROOT", "Root", "ClosFldr"
    trvFtpFldrs.Nodes("ROOT").Selected = True 'select the root node
    'Make sure we have sat the FTP Params right.
    If TxtFtp.Text <> Trim("") Or TxtUsrNm.Text <> Trim("") Then
        'Make sure the user wrote the ( ftp:// ) prefix correctly .
        If LCase(Right$(TxtFtp.Text, 6)) <> "ftp://" Then
            Beep
            LblStatus.Caption = ("Tenha certeza de que o seu servidor ftp remoto inicie com ( ftp:// ) prefix")
            Exit Sub
        End If
    'If everything is fine, then connect.
            With Inet1
                 .URL = TxtFtp.Text
                 .UserName = TxtUsrNm.Text
                 .Password = TxtPW.Text
                 Status = ("DIR") 'Triggers the DisplayData Routine
                 .Execute , "DIR"
                 UsrFriendly = ("OK!")
            End With
            Select Case Check1.Value
    Case Is = 1
                Open App.Path & "/log.txt" For Output As #Nfile
                        Print #Nfile, TxtFtp
                        Print #Nfile, TxtUsrNm
                        Print #Nfile, TxtPW
                        Print #Nfile, Check1.Value
                Close #Nfile
    Case Is = 0
                Exit Sub
End Select
    Else
            Beep
            LblStatus.Caption = ("Status : Informe o ftp remoto ftp:// + URL, UserName")
                Exit Sub
    End If
End Sub
Private Sub LblDnldFile_Click()
'Check for connection
    If Inet1.URL = Trim("") Then
        Beep
        LblStatus.Caption = ("Status : Nao Conectado")
        Exit Sub
    End If
'Create TEMP Dir to store files init (coz of the SPACES in path names are not allowed).
'You won't get any changes in the application's proccesses if you provided [Spaces] in any files or paths
'While downloading or uploading.
'[Spaces] are not allowed while working with FTP.
'Try to change C:/ with App.Path and see for yourself.
If FSO.FolderExists("C:\Temp") = False Then FSO.CreateFolder ("C:\Temp")
FtpFile = LstFtpFiles.Text
LoclFilNm = ("C:\Temp\" & FtpFile)
'Check If folder Temp contain the same file
    If FSO.FileExists("C:\Temp\" & FtpFile) Then
        LblStatus.Caption = ("Status : Arquivo " & FtpFile & " Ja existe")
        If Inet1.StillExecuting = True Then Inet1.Execute , "Fechar"
        Exit Sub
    End If
'Download :
'Open Directory, Or Download if (File) ... Selected (DBLclicked)
Status = ("GET")
Inet1.Execute , "GET " & FtpFile & " " & LoclFilNm
UsrFriendly = ("Baixado com sucesso @ (C:\Temp\)")
End Sub
Private Sub LblDelFile_Click()
'Check for connection
    If Inet1.URL = Trim("") Then
        Beep
        LblStatus.Caption = ("Status : Nao Conectado")
        Exit Sub
    End If
    'You must have the correct rights to del a file, Check your ftp:// Host Support
    If LstFtpFiles.SelCount > 0 Then
        Status = "DEL"
        Inet1.Execute , "DELETE " & Trim(LstFtpFiles.Text)
        UsrFriendly = ("Deletado com sucesso")
    Else
        LblStatus.Caption = "Nenhum arquivo para deletar, Selecione um arquivo..."
    End If
End Sub
Private Sub LblRnmFile_Click()
'Check for connection
    If Inet1.URL = Trim("") Then
        Beep
        LblStatus.Caption = ("Status : Nao Conectado")
        Exit Sub
    End If
    'rename a file
    NewName = InputBox("Informe um novo nome para o arquivo selecionado")
        If NewName = Trim("") Then
            Beep
            LblStatus.Caption = ("Não deixa o nome em branco (Novo Nome)")
            Exit Sub
        End If
    If LstFtpFiles.SelCount > 0 Then
        Status = "REN"
        Inet1.Execute , "RENAME " & Trim(LstFtpFiles.Text) & " " & Trim(NewName)
        UsrFriendly = ("Renomeado com sucesso")
    Else
        Beep
        LblStatus.Caption = "Nenhum arquivo selecionado para renomear"
    End If
End Sub
Private Sub LblUploadFile_Click()
    'place a file on the FTP server
UpldFileNm = DirLst.Path & "\" & FlsLst.FileName
    If Len(UpldFileNm) = 0 Then
        Beep
        LblStatus.Caption = ("Escolha um arquivo para enviar")
        Exit Sub
    Else
        'First we need to remove the [Spaces] from the File being Uploaded
        'Get File Name Only
        Dim N, B, Z As Integer
        Dim C, F, FtpFolder, LocalFile As String
        C = StrReverse(UpldFileNm)
        Z = InStr(C, "\")
        F = StrReverse(Left$(C, Z - 1))
        F = FSO.GetFile(F).ShortName
        UpldFileNm = FSO.GetFile(UpldFileNm).ShortPath
        FtpFolder = ("/" & LstFtpFiles.Text & F)
        Status = "PUT"
        Inet1.Execute , "PUT " & UpldFileNm & " " & FlsLst.FileName
        UsrFriendly = ("Enviado com sucesso")
    End If
End Sub
Private Sub LbNewFldr_Click()
'Check for connection
    If Inet1.URL = Trim("") Then
        Beep
        LblStatus.Caption = ("Status : Nao Conectado")
        Exit Sub
    End If
NewFldrNm = InputBox("Informe um novo nome para a pasta selecionada")
    'This is another way to check if a returned String is not Empty or a Space.
    If Len(NewFldrNm) = 0 Then
        Beep
        LblStatus.Caption = ("Não deixe o nome da pasta em branco")
        Exit Sub
    Else
        'You must have the correct rights to create a Folder, Check you Host Support
        Status = "MKDIR"
        Inet1.Execute , "MKDIR " & Trim(NewFldrNm)
        UsrFriendly = ("Nova Pasta criada com sucesso")
    End If
End Sub
Private Sub LstFtpFiles_DblClick()
'If file selected was Double-Clicked then Call Download
Call LblDnldFile_Click
End Sub
Private Sub MnuFilesExit_Click()
Dim msgB As String
    msgB = MsgBox("Confirma ?", vbYesNo + vbCritical + vbQuestion, "Sair do Programa")
        If msgB = True Then
            End
        Else
            Exit Sub
        End If
End Sub
Private Sub MnuMin_Click()
Me.WindowState = 1
End Sub
Private Sub Timer1_Timer()
'Shape1 Height resizer
    Shape1.Height = Shape1.Height + 100
        If Shape1.Height >= Me.ScaleHeight - 10 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        Exit Sub
        End If
End Sub
Private Sub Timer2_Timer()
'Shape1 Width resizer
    Shape1.Width = Shape1.Width + 10
        If Shape1.Width >= 1500 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
        Exit Sub
        End If
End Sub
Private Sub Timer3_Timer()
'Color Fader
X = X + 1
Y = Y + 2
Z = Z + 3
    If X >= 100 Then
        Y = 150
        Z = 200
            Timer3.Enabled = False
        Exit Sub
    End If
            Me.BackColor = RGB(X, Y, Z)
             'Change <BackGround Color> Property for some Controls on the form.
            Check1.BackColor = Me.BackColor
            Dim Mtxt As Control
            For Each Mtxt In Me.Controls
                If TypeOf Mtxt Is TextBox Then
                   Mtxt.BackColor = Me.BackColor
                End If
            Next
                        'This the (1st) solution (my choice) to force the user not to use the controls,
                        'untill the form loads completely, in order not to get error message ("Still Executing Last Request")
                        Dim Lbls As Control
                            For Each Lbls In Me.Controls
                                If TypeOf Lbls Is Label Then
                                    Lbls.Enabled = True
                                End If
                            Next
                        Dim trvS As Control
                            For Each trvS In Me.Controls
                                If TypeOf trvS Is TreeView Then
                                    trvS.Enabled = True
                                End If
                            Next
                        Dim FileslS As Control
                            For Each FileslS In Me.Controls
                                If TypeOf FileslS Is FileListBox Then
                                    FileslS.Enabled = True
                                End If
                            Next
End Sub
Private Sub Timer4_Timer()
Timer4.Enabled = False
Call DisplayData
End Sub
Private Sub TimWorkin_Timer()
    'Draw a red line to show that we are still executing
    If Not Inet1.StillExecuting Then
        TimWorkin.Enabled = False
        Me.MousePointer = 1
        Exit Sub
    End If
    Me.MousePointer = 11
End Sub
Private Sub trvFtpFldrs_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = "ClosFldr"
End Sub
Private Sub trvFtpFldrs_DblClick()
'The code was originally placed at (Node_Click) Event, but everytime i try to upload something
'It had to open the folder\subfolder first .
'Check for connection
    If Inet1.URL = Trim("") Then
        Beep
        LblStatus.Caption = ("Status : Nao Conectado")
        Exit Sub
    End If
    If Inet1.StillExecuting Then
        LblStatus.Caption = "Ainda trabalhando na requisição anterior ; tente novamente"
        Exit Sub
    End If
    'When we click on a node change to that dir
    'we need to build the path to the dir
    Status = "CD"
    If Trim(trvFtpFldrs.SelectedItem.Key) = "ROOT" Then 'we are at the root
    'make sure we don't have any old data in  treeview
    trvFtpFldrs.Nodes.Clear
    'we just cleared the nodes so we need to add the root node back
    'we do a DIR after the CD so we will read the data again
    'set the root node for the treeview
    trvFtpFldrs.Nodes.Add , , "ROOT", "Root", "ClosFldr"
    trvFtpFldrs.Nodes("ROOT").Selected = True 'select the root node
    Inet1.Execute , "CD /"
    Else 'not at the root
        Dim bFlag As Boolean
        Dim sTemp As String
        Dim oNode As Node
        bFlag = True
        sTemp = trvFtpFldrs.SelectedItem.Text
        Set oNode = trvFtpFldrs.SelectedItem.Parent
        If oNode.Key = "ROOT" Then 'root node
                sTemp = "/" & sTemp
        Else
            sTemp = oNode.Text & "/" & sTemp
            Do While bFlag
                Set oNode = oNode.Parent
                If oNode.Key = "ROOT" Then 'root node
                    sTemp = "/" & sTemp
                    bFlag = False
                Else
                    sTemp = oNode.Text & "/" & sTemp
                End If
            Loop
        End If
        Inet1.Execute , " CD " & sTemp
    End If

End Sub
Private Sub trvFtpFldrs_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = "OpenFldr"
End Sub
Private Sub TxtFtp_GotFocus()
TxtFtp.SelLength = Len(TxtFtp.Text)
End Sub
Private Sub TxtUsrNm_GotFocus()
TxtUsrNm.SelLength = Len(TxtUsrNm.Text)
End Sub
Private Sub TxtPW_GotFocus()
TxtPW.SelLength = Len(TxtPW.Text)
End Sub
