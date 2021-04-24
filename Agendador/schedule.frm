VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1410
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4290
   ControlBox      =   0   'False
   Icon            =   "schedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TMR 
      Interval        =   2000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Por Walter Staeblein"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agendador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   75
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agendador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EFlag As Boolean

Private Sub Form_Activate()

    If EFlag Then Unload Me
    
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        EFlag = True
    Else
        If GetSetting("AgendadorVB", "Caminho", "Agendador", "") <> App.Path Then
            SaveSetting "AgendadorVB", "Caminho", "Agendador", App.Path
        End If
        
        EFlag = (Trim(Command) = "-e")
        gHW = Me.hwnd
        Show
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unhook
    
End Sub



Private Sub TMR_Timer()

    Dim M As Byte, H As Byte, D As Byte
    Dim WD As Byte, Flag As Boolean
    Dim P As Integer, PTH As String
    
    If TMR.Interval = 2000 Then
        TMR.Interval = 60000
        Hook
        Me.Caption = "AgendadorVB"
        LoadTasks

        Me.Hide
    Else
        M = Minute(Now)
        H = Hour(Now)
        D = Day(Now)
        WD = Weekday(Now)
        
        ' Verifica se alguma tarefa deve ser rodada
        For I = 0 To UBound(Tasks, 1)
        
            ' Checa a data inicial
            If Tasks(I).DataIni >= Date Then
                Select Case Tasks(I).Frequencia
                    Case 2
                    ' Diario
                    If H = Tasks(I).Hora And M = Tasks(I).Minuto Then Flag = True
                    
                    Case 3
                    ' Semanal
                    If WD = Tasks(I).Dia And H = Tasks(I).Hora And M = Tasks(I).Minuto Then Flag = True
                    
                    Case 4
                    ' Mensal
                    If D = Tasks(I).Dia And H = Tasks(I).Hora And M = Tasks(I).Minuto Then Flag = True
                End Select
            End If
            
            If Flag = True Then
                ' Separa o caminho do nome de arquivo
                P = InStrRev(Tasks(I).Arquivo, "\")
                If P > 0 Then
                    PTH = Left$(Tasks(I).Arquivo, P)
                Else
                    PTH = "C:\"
                End If
                
                ' Executa o arquivo
                Resp = ShellExecute(Me.hwnd, "Open", Tasks(I).Arquivo, vbNullString, PTH, SW_SHOWNORMAL)
            End If
        
        Next
        
    End If
    
End Sub

Sub LoadTasks()

    Dim A As Variant, N As Integer
    Dim I As Integer, TMP As Variant
    
    A = GetAllSettings("AgendadorVB", "Tarefas")
    If IsArray(A) = True Then
        N = UBound(A, 1)
        ReDim Tasks(N) As Sched
        
        For I = 0 To N
            Nome = A(I, 0)
            TMP = Split(A(I, 1), Chr(160))
            
            ' Critica
            If TMP(1) < 1 Or TMP(1) > 4 Then TMP(1) = 1
            If IsNumeric(TMP(2)) = False Then TMP(2) = 1
            If IsNumeric(TMP(3)) = False Then TMP(3) = 12
            If IsNumeric(TMP(4)) = False Then TMP(4) = 0
            
            ' Armazena
            Tasks(I).Nome = Nome
            Tasks(I).Arquivo = TMP(0)
            Tasks(I).Frequencia = TMP(1)
            Tasks(I).Dia = TMP(2)
            Tasks(I).Hora = TMP(3)
            Tasks(I).Minuto = TMP(4)
            Tasks(I).DataIni = TMP(5)
            Erase TMP
        Next
    End If
    
End Sub

Sub Fim()

    Unload Me
    
End Sub
