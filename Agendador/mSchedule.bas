Attribute VB_Name = "mSchedule"
' ------
' SERVER
' ------
Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Enum Tipos
    Manual = 1
    Diario = 2
    Semanal = 3
    Mensal = 4
End Enum

Public Type Sched
    Nome As String
    Arquivo As String
    Frequencia As Tipos
    Dia As Byte
    Hora As Byte
    Minuto As Byte
    DataIni As Date
End Type

Public Tasks() As Sched
Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Public lpPrevWndProc As Long
Public gHW As Long

'Copies a block of memory from one location to another.
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
    If uMsg = WM_COPYDATA Then
        Call InterProcessComms(lngParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lngParam)
End Function


Sub InterProcessComms(lngParam As Long)

          Dim cdCopyData As COPYDATASTRUCT
          Dim byteBuffer(1 To 255) As Byte
          Dim strTemp As String
          
          Call CopyMemory(cdCopyData, ByVal lngParam, Len(cdCopyData))

          Select Case cdCopyData.dwData
            Case 3
                    Call CopyMemory(byteBuffer(1), ByVal cdCopyData.lpData, cdCopyData.cbData)
                    strTemp = StrConv(byteBuffer, vbUnicode)
                    strTemp = Left$(strTemp, InStr(1, strTemp, Chr$(0)) - 1)
                    RunTask strTemp
          End Select
End Sub

Private Sub RunTask(ByVal BUF As String)

    Dim CMD As String, Nome As String, PTH As String
    Dim P As Long, Resp As Long, Arq As String
    
    CMD = Left(BUF, 3)
    Nome = Right(BUF, Len(BUF) - 3)
    
    Select Case CMD
        Case "EXE"
            ' Executa tarefa agora
            For I = 0 To UBound(Tasks, 1)
                If LCase(Nome) = LCase(Tasks(I).Nome) Then
                    
                    ' Separa o caminho do nome de arquivo
                    P = InStrRev(Tasks(I).Arquivo, "\")
                    If P > 0 Then
                        PTH = Left$(Tasks(I).Arquivo, ppos)
                    Else
                        PTH = "C:\"
                    End If

                    Resp = ShellExecute(Form1.hwnd, "open", Tasks(I).Arquivo, vbNullString, PTH, SW_SHOWNORMAL)
                    SaveSetting "AgendadorVB", "Tarefas", Tasks(I).Nome, Tasks(I).Arquivo & Chr(160) & Tasks(I).Frequencia & Chr(160) & Tasks(I).Dia & Chr(160) & Tasks(I).Hora & Chr(160) & Tasks(I).Minuto & Chr(160) & Tasks(I).DataIni & Chr(160) & Now
                    Exit For
                End If
            Next
            
        Case "REF"
            ' Atualiza lista de tarefas
            Form1.LoadTasks
        
        Case "END"
            Form1.Fim
            
        Case "ADD"
            ' Adiciona a si mesmo no start automático do computador
            Arq = LCase(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName)
            If Right(Arq, 4) <> ".exe" Then Arq = Arq & ".exe"
            AddToRun "AgendadorVB", Arq
            
        Case "REM"
            ' Remove a si mesmo do start automático do computador
            RemoveFromRun "AgendadorVB"
            
    End Select
            
End Sub
