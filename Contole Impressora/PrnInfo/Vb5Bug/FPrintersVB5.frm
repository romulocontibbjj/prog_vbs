VERSION 5.00
Begin VB.Form FPrinters 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test for Q253612 problems"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   435
      Left            =   1620
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4275
   End
End
Attribute VB_Name = "FPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_prns As CPrinters

Private Sub m_prns_PrinterAdded(ByVal DeviceName As String)
   List1.AddItem DeviceName
End Sub

Private Sub Command1_Click()
   List1.Clear
   If m_prns Is Nothing Then
      Set m_prns = New CPrinters
   Else
      m_prns.Refresh
   End If
   If m_prns.PrintersCollectionBad Then
      MsgBox "If you are running a VB5 application on this" & vbCrLf & _
             "machine, the Printers collection is incomplete." & vbCrLf & vbCrLf & _
             "See Microsoft KB article Q253612 for details.", vbCritical, _
             "Printers Collection Hosed in VB5"
   End If
End Sub

