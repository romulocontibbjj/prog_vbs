VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4215
   FillStyle       =   0  'Solid
   Height          =   4215
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   Top             =   1170
   Width           =   4335
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265

Public Function CompassToRadians(degrees As Single) As Single
    CompassToRadians = PI / 2 - PI * degrees / 180
End Function

' ********************************************
' Draw a pie slice clockwise between the two
' angles. Measure angles in compass degrees.
' ********************************************
Public Sub PieSlice(obj As Object, ByVal X As Single, ByVal Y As Single, ByVal radius As Single, ByVal color As Long, ByVal start_ang As Single, ByVal end_ang As Single)
Dim ang1 As Single
Dim ang2 As Single
Dim tmp As Single

    ang1 = CompassToRadians(end_ang)
    ang2 = CompassToRadians(start_ang)

    ' Zero is bad.
    If ang1 = 0 Then ang1 = 0.01
    If ang2 = 0 Then ang2 = 0.01
    
    ' Make ang2 bigger.
    If ang1 > ang2 Then
        tmp = ang1
        ang1 = ang2
        ang2 = tmp
    End If
    
    If ang2 < 0 Then ' Make ang2 > 0.
        ang1 = ang1 + 2 * PI
        ang2 = ang2 + 2 * PI
    End If
    
    If ang1 <= 0 Then
        obj.Circle (X, Y), radius, obj.FillColor, -(2 * PI + ang1), -0.01
        obj.Circle (X, Y), radius, obj.FillColor, -0.01, -ang2
        
        obj.Circle (X, Y), radius, color, -(2 * PI + ang1), 0.01
        obj.Circle (X, Y), radius, color, 0.01, -ang2
    Else
        obj.Circle (X, Y), radius, color, -ang1, -ang2
    End If
End Sub


Private Sub Command1_Click()
Dim ang1 As Integer
Dim ang2 As Integer
Static i As Integer

    Cls

    ang2 = 0
    Do While ang2 < 360
        ang1 = ang2
        ang2 = ang1 + Rnd * 60 + 20
        FillColor = QBColor(i Mod 16)
        i = i + 1
        PieSlice Me, 150, 150, 90, vbBlack, ang1, ang2
    Loop
End Sub

