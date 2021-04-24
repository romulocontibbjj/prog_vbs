Attribute VB_Name = "Delegator"
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Declare Function GlobalLock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal HMEM As Long) As Long


Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40


Public Type FunctionSPointerS
FunctionPtr As Long
FunctionAddress As Long
End Type

Public Function CalculateSpaceForDelegation(ByVal NumberOfParameters As Byte) As Long
CalculateSpaceForDelegation = 31 + NumberOfParameters * 3
End Function

Public Function DelegateFunction(ByVal CallingADR As Long, Obj As Object, ByVal MethodAddress As Long, ByVal NumberOfParameters As Byte) As Boolean
On Error GoTo NotSuccess
Dim TmpA As Long
TmpA = CallingADR
CopyMemory ByVal CallingADR, &H68EC8B55, 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
CallingADR = CallingADR + 4

Dim StackP As Byte
StackP = 4 + 4 * NumberOfParameters

For u = 1 To NumberOfParameters
CopyMemory ByVal CallingADR, CInt(&H75FF), 2
CallingADR = CallingADR + 2
CopyMemory ByVal CallingADR, StackP, 1
CallingADR = CallingADR + 1
StackP = StackP - 4
Next u

CopyMemory ByVal CallingADR, CByte(&H68), 1
CallingADR = CallingADR + 1
CopyMemory ByVal CallingADR, ObjPtr(Obj), 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, CByte(&HE8), 1
CallingADR = CallingADR + 1
Dim PERFCALL As Long
PERFCALL = CallingADR - TmpA - 1
PERFCALL = MethodAddress - (TmpA + (CallingADR - TmpA - 1)) - 5
CopyMemory ByVal CallingADR, PERFCALL, 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, CByte(&HA1), 1
CallingADR = CallingADR + 1
CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, CInt(&HC2C9), 2

CallingADR = CallingADR + 2
CopyMemory ByVal CallingADR, CInt(NumberOfParameters * 4), 2

'FINALLY !!! ABSOLUTE CALLING RUTINE!


'WHAT IS BEHIND ASM CODE:
'*****************************
'PUSH EBP
'MOV EBP,ESP
'PUSH OFFSET RETURN ADDRESS

'*********** Depend on Number of Parameters
'PUSH EBP+XX
'  .......
'PUSH EBP+10
'PUSH EBP+0C
'PUSH EBP+08
'***********

'PUSH OBJECT POINTER
'CALL POINTER OBJECT.METHOD
'MOV EAX,DWORD PTR [OFFSET RETURN ADDRESS]
'LEAVE
'RET 00XX Depend on Number of Parameters
'TEMPSTORE dd 00 <------RETURN ADDRESS PTR

'Thats IT! Nothing less than 39 BYTES Of ASM Code!

DelegateFunction = True
Exit Function
NotSuccess:
On Error GoTo 0
End Function
Public Function GetObjectFunctionsPointers(Obj As Object, ByVal NumberOfMethods As Long, Optional ByVal PublicVarNumber As Long, Optional ByVal PublicObjVariantNumber As Long) As FunctionSPointerS()
Dim FPS() As FunctionSPointerS
ReDim FPS(NumberOfMethods - 1)
Dim OBJ1 As Long
OBJ1 = ObjPtr(Obj)
Dim VTable As Long
CopyMemory VTable, ByVal OBJ1, 4
Dim PTX As Long
Dim u As Long
For u = 0 To NumberOfMethods - 1
PTX = VTable + 28 + (PublicVarNumber * 2 * 4) + (PublicObjVariantNumber * 3 * 4) + u * 4
CopyMemory FPS(u).FunctionPtr, PTX, 4
CopyMemory FPS(u).FunctionAddress, ByVal PTX, 4
Next u
GetObjectFunctionsPointers = FPS
End Function



