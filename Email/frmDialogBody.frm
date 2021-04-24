VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDialog_Body 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Page Body"
   ClientHeight    =   4185
   ClientLeft      =   4170
   ClientTop       =   3960
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDialogBody.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   480
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   6975
      Begin VB.Frame Preview 
         Caption         =   "Preview"
         Height          =   2415
         Left            =   3480
         TabIndex        =   19
         Top             =   375
         Width           =   3375
         Begin VB.PictureBox Layout 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   3075
            TabIndex        =   20
            Top             =   240
            Width           =   3135
            Begin VB.Label ActiveLink 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "This is an Active Link"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   1680
               Width           =   1800
            End
            Begin VB.Label VisitedLink 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "This is a Visited Link"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   1200
               Width           =   1740
            End
            Begin VB.Label Link 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "This is a Link"
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   1110
            End
            Begin VB.Label Regular 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "This is Regular Text"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   1695
            End
         End
      End
      Begin VB.TextBox Combo5 
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
         Left            =   1830
         TabIndex        =   18
         Top             =   2295
         Width           =   1245
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   2295
         Width           =   255
      End
      Begin VB.TextBox Combo4 
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
         Left            =   1830
         TabIndex        =   16
         Top             =   1815
         Width           =   1245
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1815
         Width           =   255
      End
      Begin VB.TextBox Combo3 
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
         Left            =   1830
         TabIndex        =   14
         Top             =   1335
         Width           =   1245
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   1335
         Width           =   255
      End
      Begin VB.TextBox Combo2 
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
         Left            =   1830
         TabIndex        =   12
         Top             =   855
         Width           =   1245
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   855
         Width           =   255
      End
      Begin VB.TextBox Combo1 
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
         Left            =   1830
         TabIndex        =   10
         Top             =   375
         Width           =   1245
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   375
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Active Link Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2295
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Visited Link Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1815
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Link Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1335
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Text Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   855
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Background Color:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   375
         Width           =   1620
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Finish"
      Height          =   375
      Left            =   4740
      TabIndex        =   2
      Top             =   3705
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6090
      TabIndex        =   1
      Top             =   3705
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3510
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6191
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Insert Page Body "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find"
      Height          =   285
      Left            =   6300
      TabIndex        =   25
      Top             =   690
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   26
      Top             =   690
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background Image:"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   690
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "frmDialog_Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Respond to the buttons with these below
' ------------------------------------------------------------

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next

    ' Strings for HTML Body
    Dim BackImage2 As String
    Dim BackGroundColor2 As String
    Dim TextColor2 As String
    Dim LinkColor2 As String
    Dim VColor2 As String
    Dim Acolor2 As String

    ' Put information from form into string variables
    BackImage2 = Text1.Text
    BackGroundColor2 = Combo1.Text
    TextColor2 = combo2.Text
    LinkColor2 = combo3.Text
    VColor2 = combo4.Text
    Acolor2 = Combo5.Text
    
    ' Unload Wizard Screen
    Unload Me
    
    If BackGroundColor2 = "" Then
    Else
        ColorBody = BackGroundColor2
    End If
    
    If TextColor2 = "" Then
    Else
        ColorText = TextColor2
    End If
    
    If LinkColor2 = "" Then
    Else
        ColorLink = LinkColor2
    End If
    
    If VColor2 = "" Then
    Else
        ColorVisited = VColor2
    End If
    
    If Acolor2 = "" Then
    Else
        ColorActive = Acolor2
    End If
        
End Sub

' Grab the colors for the color dialogs
' -------------------------------------------------------

Private Sub Command4_Click()
    cmDialog.flags = cdlCCFullOpen Or cdlCCRGBInit
    'cmdialog.color = Command1.BackColor
    cmDialog.ShowColor

    Layout.BackColor = cmDialog.color
    
    Dim r As Integer
    Dim G As Integer
    Dim B As Integer
    
    ExtractRGB cmDialog.color, r, G, B
    Combo1.Text = "#" & Format(Hex(r), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")
End Sub

Private Sub Command5_Click()
    cmDialog.flags = cdlCCFullOpen Or cdlCCRGBInit
    'cmdialog.color = Command1.BackColor
    cmDialog.ShowColor

    Regular.ForeColor = cmDialog.color
    
    Dim r As Integer
    Dim G As Integer
    Dim B As Integer
    
    ExtractRGB cmDialog.color, r, G, B
    combo2.Text = "#" & Format(Hex(r), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")
End Sub

Private Sub Command6_Click()
    cmDialog.flags = cdlCCFullOpen Or cdlCCRGBInit
    'cmdialog.color = Command1.BackColor
    cmDialog.ShowColor

    Link.ForeColor = cmDialog.color
    
    Dim r As Integer
    Dim G As Integer
    Dim B As Integer
    
    ExtractRGB cmDialog.color, r, G, B
    combo3.Text = "#" & Format(Hex(r), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")
End Sub

Private Sub Command7_Click()
    cmDialog.flags = cdlCCFullOpen Or cdlCCRGBInit
    'cmdialog.color = Command1.BackColor
    cmDialog.ShowColor

    VisitedLink.ForeColor = cmDialog.color

    Dim r As Integer
    Dim G As Integer
    Dim B As Integer
    
    ExtractRGB cmDialog.color, r, G, B
    combo4.Text = "#" & Format(Hex(r), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")
End Sub

Private Sub Command8_Click()
    cmDialog.flags = cdlCCFullOpen Or cdlCCRGBInit
    'cmdialog.color = Command1.BackColor
    cmDialog.ShowColor

    ActiveLink.ForeColor = cmDialog.color

    Dim r As Integer
    Dim G As Integer
    Dim B As Integer
    
    ExtractRGB cmDialog.color, r, G, B
    Combo5.Text = "#" & Format(Hex(r), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")
End Sub

