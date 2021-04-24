Attribute VB_Name = "Module1"
Public Function Xminutes(xhora As String)
Dim xhora1 As Integer
Dim xminute As Integer


xhora1 = Int(Mid(xhora, 1, 2)) * 60
xminute = Int(Mid(xhora, 4, 2))


Xminutes = (xhora1 + xminute)


End Function
