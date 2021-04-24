VERSION 5.00
Begin VB.Form BIsseccao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BISSECÇÃO"
   ClientHeight    =   2970
   ClientLeft      =   4710
   ClientTop       =   5700
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2835
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5655
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1995
         Left            =   2880
         TabIndex        =   8
         Top             =   420
         Width           =   2355
         Begin VB.Label Label7 
            Caption         =   "E:"
            Height          =   315
            Left            =   540
            TabIndex        =   17
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label lab_k 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   780
            TabIndex        =   16
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "K:"
            Height          =   255
            Left            =   540
            TabIndex        =   15
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "B:"
            Height          =   255
            Left            =   540
            TabIndex        =   14
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "A:"
            Height          =   315
            Left            =   540
            TabIndex        =   13
            Top             =   360
            Width           =   195
         End
         Begin VB.Label lab_a 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   780
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lab_b 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   780
            TabIndex        =   11
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lab_e 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   780
            TabIndex        =   10
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ENTRADA DE DADOS"
         Height          =   1995
         Left            =   540
         TabIndex        =   1
         Top             =   420
         Width           =   2355
         Begin VB.TextBox txt_a 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   600
            TabIndex        =   5
            Top             =   360
            Width           =   1395
         End
         Begin VB.TextBox txt_b 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   600
            TabIndex        =   4
            Top             =   720
            Width           =   1395
         End
         Begin VB.TextBox txt_E 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   600
            TabIndex        =   3
            Top             =   1080
            Width           =   1395
         End
         Begin VB.CommandButton cmd_calc 
            Caption         =   "Calcula"
            Height          =   255
            Left            =   600
            TabIndex        =   2
            Top             =   1440
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "E:"
            Height          =   315
            Left            =   360
            TabIndex        =   9
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "A:"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "B:"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "BIsseccao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_calc_Click()
Dim i As Integer
Dim xa As Double
Dim xb As Double
Dim xe As Double
Dim xk As Double
Dim ra As Double
Dim rb As Double
Dim rm As Double
Dim xm As Double
Dim xe1 As Integer
Dim xa_b As Integer
Dim xe_1 As Double
Dim xreq As Double
Dim x1 As Integer

xa = txt_a.Text
xb = txt_b.Text
xe = txt_E.Text
xm = (xa + xb) / 2

xa_b = xb - xa
xe_1 = xa_b / xe


xk = 3


For i = 0 To xk Step 1
'ra = (1.8 * Exp(-(xa))) - 2 * Exp(-(xa) / 2)
ra = ((xa) ^ 3) - 10
'rb = (1.8 * Exp(-(xb))) - 2 * Exp(-(xb) / 2)
rb = ((xb) ^ 3) - 10
xm = (xa + xb) / 2
rm = ((xm) ^ 3) - 10
'rm = (1.8 * Exp(-(xm))) - 2 * Exp(-(xm) / 2)

'rm = (xm ^ 3) - 10

If xreq = xb - xa < 0 Then
    Exit For
End If


'If xreq < xe Then
        'ra = (1.8 * Exp(-(xa))) - 2 * Exp(-(xa) / 2)
ra = ((xa) ^ 3) - 10
'rb = (1.8 * Exp(-(xb))) - 2 * Exp(-(xb) / 2)
rb = ((xb) ^ 3) - 10
'xm = (xa + xb) / 2
rm = ((xm) ^ 3) - 10
'rm = (1.8 * Exp(-(xm))) - 2 * Exp(-(xm) / 2)

    'Exit For
             
'Else
If (ra * rm) < 0 Then
    xb = xm
ElseIf (rb * rm) < 0 Then
    xa = xm
End If
'End If

If xreq = xb - xa < 0 Then
    Exit For
End If



Next
lab_k.Caption = i - 1
lab_a.Caption = xa
lab_b.Caption = xb
lab_e.Caption = xreq



End Sub
