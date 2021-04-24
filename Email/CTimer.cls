VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CElapsedTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Elapsed time class
' Copyright (c) 1997 by Desaware Inc. All Rights Reserved

Option Explicit

Private Declare Function GetTickCount& Lib "kernel32" ()

Private m_CreationTime As Long
Private m_StopTime As Long


' Update the creation time. This should always
' be called because class initialization is not
' as controllable.
Public Sub StartTheClock()
    m_CreationTime = GetTickCount()
End Sub

' Mark the stop time. This is called automatically
' the first time you request the elapsed time for an
' object.
Public Sub StopTheClock()
    m_StopTime = GetTickCount()
End Sub

' Get a formatted string for the time in microseconds
Public Function Elapsed(Optional ByVal repetitions As Long = 1) As String
    Dim timeval As Long
    If m_StopTime = 0 Then StopTheClock
    timeval = m_StopTime - m_CreationTime
    ' timeval <0 indicates StartTheClock was never called
    ' You could raise an error here instead
    If timeval < 0 Then timeval = 0
    ' timeval is the difference in milliseconds
    Elapsed = Format$(CDbl(timeval) / repetitions, "0.###")
End Function

