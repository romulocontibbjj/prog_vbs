VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2445
   ClientLeft      =   5100
   ClientTop       =   5985
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4155
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim xarq As String

xarq = "C:\BD\bd1.mdb"

If Xbusca(UCase(xarq), True) = True Then
    
    MsgBox "Arquivo já No Computador" & Chr$(13) & xarq
    
Else
    MkDir ("C:\BD")
    FileCopy "D:\bd1.mdb", "C:\BD\" & "bd1.mdb"

End If

Unload Me

End Sub
