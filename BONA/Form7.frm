VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5145
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   11310
   LinkTopic       =   "Form7"
   ScaleHeight     =   5145
   ScaleWidth      =   11310
   Begin BONAGURA.isButton isButton2 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "Form7.frx":0000
      Style           =   0
      Caption         =   "isButton2"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Command1"
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin BONAGURA.isButton isButton1 
      Height          =   540
      Left            =   5400
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   953
      Icon            =   "Form7.frx":001C
      Style           =   0
      Caption         =   "isButton1"
      IconAlign       =   1
      iNonThemeStyle  =   0
      BackColor       =   33023
      HighlightColor  =   12582912
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
isButton1.Style = [Mac OSX]
isButton2.Style = [Mac OSX]

isButton1.ToolTipIcon = TTIconInfo
isButton1.ToolTipType = TTBALLONN

isButton1.FontHighlightColor = &HA0FF&



  
        
End Sub

