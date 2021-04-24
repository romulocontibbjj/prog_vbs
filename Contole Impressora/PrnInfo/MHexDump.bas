Attribute VB_Name = "MHexDump"
' *************************************************************************
'  Copyright ©2000 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'Public Sub Main()
'   Dim s As String
'   Dim b() As Byte
'   s = "abcdef"
'   Debug.Print HexDump(StrPtr(s) - 4, 16)
'   Debug.Print HexDump(StrPtr(StrConv(s, vbFromUnicode)) - 4, 16)
'   b = s
'   Debug.Print HexDump(VarPtr(b(0)) - 4, 16)
'End Sub

Public Function FmtHex(ByVal InVal As Long, ByVal OutLen As Integer) As String
   ' Left pad with zeros to OutLen.
   FmtHex = Right$(String$(OutLen, "0") & Hex$(InVal), OutLen)
End Function

Public Function HexDump(ByVal lpBuffer As Long, ByVal nBytes As Long) As String
   Dim i As Long, j As Long
   Dim ba() As Byte
   Dim sRet As String
   Dim dBytes As Long
   
   ' Size recieving buffer as requested,
   ' then sling memory block to buffer.
   ReDim ba(0 To nBytes - 1) As Byte
   Call CopyMemory(ba(0), ByVal lpBuffer, nBytes)
   sRet = String(81, "=") & vbCrLf & _
          "lpBuffer = &h" & Hex$(lpBuffer) & _
          "   nBytes = " & nBytes
   
   ' Buffer may well not be even multiple of 16.
   ' If not, we need to round up.
   If nBytes Mod 16 = 0 Then
      dBytes = nBytes
   Else
      dBytes = ((nBytes \ 16) + 1) * 16
   End If
   
   ' Loop through buffer, displaying 16 bytes per
   ' row. Preface with offset, trail with ASCII.
   For i = 0 To (dBytes - 1)
      ' Add address and offset from beginning
      ' if at the start of new row.
      If (i Mod 16) = 0 Then
         sRet = sRet & vbCrLf & Right$("00000000" _
                & Hex$(lpBuffer + i), 8) & "  " & _
                Right$("0000" & Hex$(i), 4) & "  "
      End If
      
      ' Append this byte.
      If i < nBytes Then
         sRet = sRet & Right$("0" & Hex(ba(i)), 2)
      Else
         sRet = sRet & "  "
      End If
      
      ' Special handling...
      If (i Mod 16) = 15 Then
         ' Display last 16 characters in
         ' ASCII if at end of row.
         sRet = sRet & "  "
         For j = (i - 15) To i
            If j < nBytes Then
               If ba(j) >= 32 And ba(j) <= 126 Then
                  sRet = sRet & Chr$(ba(j))
               Else
                  sRet = sRet & "."
               End If
            End If
         Next j
      ElseIf (i Mod 8) = 7 Then
         ' Insert hyphen between 8th and
         ' 9th bytes of hex display.
         sRet = sRet & "-"
      Else
         ' Insert space between other bytes.
         sRet = sRet & " "
      End If
   Next i
   HexDump = sRet & vbCrLf & String(81, "=") & vbCrLf
End Function

Public Function PointerToStringA(ByVal lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function

Public Function PointerToStringW(ByVal lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long

   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

Public Function PointerToDWord(ByVal lpDWord As Long) As Long
   Dim nRet As Long
   If lpDWord Then
      CopyMemory nRet, ByVal lpDWord, 4
      PointerToDWord = nRet
   End If
End Function

Public Function LoWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(LoWord, LongIn, 2)
End Function

Public Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function


