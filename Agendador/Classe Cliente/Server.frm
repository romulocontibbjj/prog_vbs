VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo do Agendador VB"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Ver"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "A"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Adiciona"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Executa"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ag As New cAgendador

Private Sub Command1_Click()

   If CheckRunning Then Ag.Execute List1.Text
    
End Sub

Private Sub Command2_Click()

    If CheckRunning Then Ag.Terminate
    
End Sub

Private Sub Command3_Click()

    Ag.Remove List1.Text
    Form1.LoadList

End Sub

Private Sub Command4_Click()

    Form2.Show 1
    
End Sub

Private Sub Command5_Click()

    If CheckRunning Then Ag.Add2StartUp
    
End Sub

Private Sub Command6_Click()

    Dim BUF As String, D As String
    
    If List1.Text <> "" Then
        Ag.TaskName = List1.Text
        BUF = "Nome: " & List1.Text & vbCrLf
        BUF = BUF & "Arquivo: " & Ag.Arquivo & vbCrLf
        BUF = BUF & "Frequencia: "
        Select Case Ag.Frequencia
            Case 1
                BUF = BUF & "Manual"
                D = "N/D"
            Case 2
                BUF = BUF & "Diario"
                D = "N/D"
            Case 3
                BUF = BUF & "Semanal"
                D = Format(Ag.Dia, "dddd")
            Case 4
                BUF = BUF & "Mensal"
                D = Ag.Dia
        End Select
                
        BUF = BUF & vbCrLf & "Data Inicial: " & Ag.StartDate & vbCrLf
        BUF = BUF & "Último Processamento: " & Ag.LastRun & vbCrLf
        BUF = BUF & "Dia: " & D & vbCrLf
        BUF = BUF & "Hora: " & Ag.Hora & ":" & Ag.Minuto & vbCrLf

        MsgBox BUF, vbInformation
    End If
    
End Sub

Private Sub Form_Load()

    Ag.WinHandle = Me.hwnd
    LoadList
    
End Sub

Sub LoadList()

    List1.Clear
    
    X = Ag.List
    
    If IsArray(X) Then
        For i = 0 To UBound(X)
            List1.AddItem X(i)
        Next
    End If

End Sub

Function CheckRunning() As Boolean

    If Ag.IsRunning = False Then
        MsgBox "Agendador não está rodando", vbCritical
    Else
        CheckRunning = True
    End If

End Function
