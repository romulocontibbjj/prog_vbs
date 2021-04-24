VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FMR_varios 
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   1380
   ClientTop       =   1545
   ClientWidth     =   10065
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   10065
   Begin VB.Frame Frame2 
      Caption         =   "Concatenar"
      Height          =   2295
      Left            =   4320
      TabIndex        =   8
      Top             =   240
      Width           =   5655
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Text            =   "C"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Text            =   "D"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Text            =   "E"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   "10"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Text            =   "10"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Text            =   "10"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_exe 
      Caption         =   "EXE"
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmd_sons 
      Caption         =   "SONS"
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   5775
      Begin MSMask.MaskEdBox MaskEdBox2 
         Bindings        =   "FMR_varios.frx":0000
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Cmd_diasUteis 
         Caption         =   "Dias Úteis"
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmd_winzip 
      Caption         =   "WINZIP"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton EXPLORER 
      Caption         =   "INTERNET EXPLORER"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "FMR_varios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_diasUteis_Click()
Dim xdiasuteis As Integer

xdiasuteis = DiasUteis(CDate(MaskEdBox1), CDate(MaskEdBox2))

MsgBox xdiasuteis



End Sub




Private Sub cmd_exe_Click()


retval = Shell("c:\Kazaa2005.Exe", vbNormalFocus) ' Run nome do exe(exp. ;vsb).


End Sub

Private Sub cmd_winzip_Click()
Dim Compri
Dim fecha
Dim Archivo

fecha = Date$

Archivo = "ABBOTT.XLS"

Compri = Shell("C:\Arquivos de programas\WinZip\WINZIP32.EXE -a -en c:\" & Archivo & "")

MsgBox "OK"

End Sub



Private Sub Command1_Click()

Label1.Caption = ConcatenaString(Text1.Text, Int(Text4.Text), Text10.Text, Text7.Text)
Label2.Caption = ConcatenaString(Text2.Text, Int(Text5.Text), Text11.Text, Text8.Text)
Label3.Caption = ConcatenaString(Text3.Text, Int(Text6.Text), Text12.Text, Text9.Text)

End Sub

Private Sub EXPLORER_Click()

Shell "C:\Arquivos de programas\Internet Explorer\IEXPLORE.EXE  http://pdj.juco.zip.net"

End Sub
