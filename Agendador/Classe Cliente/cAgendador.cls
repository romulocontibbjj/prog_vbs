VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAgendador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------
' CLIENTE
' -------

Private Const WM_COPYDATA = &H4A
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Enum Tipos
    Uma_Vez = 1
    Diario = 2
    Semanal = 3
    Mensal = 4
End Enum

Private bCanSend As Boolean
Private iDia As Byte
Private iHora As Byte
Private iMinuto As Byte
Private iFrequency As Byte
Private sArquivo As String
Private sNome As String
Private dLastRun As Date
Private dStartDate As Date

Private Sub Send(ByVal BUF As String)

    Dim cdCopyData As COPYDATASTRUCT
    Dim ThWnd As Long
    Dim byteBuffer(1 To 255) As Byte
    
    ' Get the hWnd of the target application
    ThWnd = FindWindow(vbNullString, "AgendadorVB")
    
        ' Copy the string into a byte array, converting it to ASCII
    Call CopyMemory(byteBuffer(1), ByVal BUF, Len(BUF))
    cdCopyData.dwData = 3
    cdCopyData.cbData = Len(BUF) + 1
    cdCopyData.lpData = VarPtr(byteBuffer(1))
    i = SendMessage(ThWnd, WM_COPYDATA, lWinHandle, cdCopyData)
        
End Sub

Public Sub Add2StartUp(Optional DoIt As Boolean = True)

    If DoIt = True Then
        Send "ADD"
    Else
        Send "REM"
    End If
    
End Sub

Public Sub Save(Nome As String)

    ' Salva tarefa no agendador
    SaveSetting "AgendadorVB", "Tarefas", Nome, sArquivo & Chr(160) & CStr(iFrequency) & Chr(160) & CStr(iDia) & Chr(160) & CStr(iHora) & Chr(160) & CStr(iMinuto) & Chr(160) & CStr(dStartDate) & Chr(160) & "0"
    ' Manda o agendador ler
    Send "REF" & Nome
    
End Sub

Public Sub Execute(Nome As String)
    
    ' Executa tarefa agora
    Send "EXE" & Nome
 
End Sub

Public Sub Terminate()

    Send "END"
    
End Sub

Public Function List() As Variant

    ' Retorna array com os nomes das tarefas

    On Error Resume Next
    
    Dim TMP As Variant, BUF As String
    
    TMP = GetAllSettings("AgendadorVB", "Tarefas")
    For i = 0 To UBound(TMP, 1)
        BUF = BUF & TMP(i, 0) & vbCrLf
    Next
    If Right(BUF, 2) = vbCrLf Then BUF = Left(BUF, Len(BUF) - 2)
    
    List = Split(BUF, vbCrLf)
    
End Function

Public Function Remove(ByVal Nome As String)

    Err.Clear
    On Error Resume Next
    
    ' Remove uma tarefa do agendador
    DeleteSetting "AgendadorVB", "Tarefas", Nome
    
    Remove = (Err.Number = 0)
    
    If Remove = 0 Then
        ' Manda o agendador ler
        Send "REF" & Nome
    End If
    
End Function

Public Property Get IsRunning()

      IsRunning = (FindWindow(vbNullString, "AgendadorVB") <> 0)
      
End Property


Public Property Get Frequencia() As Tipos

      Frequencia = iFrequency
      
End Property

Public Property Let Frequencia(iArg As Tipos)

      iFrequency = iArg
      
End Property

Public Property Get Dia() As Byte

      Dia = iDia
      
End Property

Public Property Let Dia(iArg As Byte)

    Select Case iFrequency
    Case 3
        If iArg > 7 Then iArg = 7
        If iArg < 1 Then iArg = 1
        
    Case 4
        If iArg > 31 Then iArg = 31
        If iArg < 1 Then iArg = 1
    End Select
    
    iDia = iArg
      
End Property

Public Property Get Hora() As Byte

      Hora = iHora
      
End Property

Public Property Let Hora(iArg As Byte)

      iHora = iArg
      
End Property

Public Property Get Minuto() As Byte

      Minuto = iMinuto
      
End Property

Public Property Let Minuto(iArg As Byte)

    iMinuto = iArg
      
End Property

Public Property Get Arquivo() As String

    Arquivo = sArquivo
      
End Property

Public Property Let Arquivo(sArg As String)

    sArquivo = sArg
      
End Property

Public Property Get TaskName() As String

    TaskName = sNome
      
End Property

Public Property Get WinHandle() As Long

      WinHandle = lWinHandle
      
End Property

Public Property Let WinHandle(lArg As Long)

      lWinHandle = lArg
      
End Property

Public Property Get LastRun() As Date

      LastRun = dLastRun
      
End Property

Public Property Get StartDate() As Date

      StartDate = dStartDate
      
End Property

Public Property Let StartDate(dArg As Date)

      dStartDate = CDate(CLng(dArg))
      
End Property

Public Property Let TaskName(sArg As String)

    ' Busca o item cujo nome está em sArg no registro e preenche as propriedades
    
    Dim TMP As Variant
    
    sNome = sArg
    TMP = Split(GetSetting("AgendadorVB", "Tarefas", sNome, ""), Chr(160))
    sArquivo = TMP(0)
    iFrequency = CByte(TMP(1))
    iDia = CByte(TMP(2))
    iHora = CByte(TMP(3))
    iMinuto = CByte(TMP(4))
    dStartDate = CDate(TMP(5))
    dLastRun = CDate(TMP(6))
    
End Property

Private Sub Class_Initialize()

    bCanSend = IsNumeric(FindWindow(vbNullString, "AgendadorVB"))
    iFrequency = 1
    iDia = Day(Now)
    iHora = Hour(Now)
    iMinuto = Minute(Now)
    sArquivo = ""
    
End Sub
