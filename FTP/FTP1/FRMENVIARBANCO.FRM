VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmEnviarBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualizar na Internet"
   ClientHeight    =   3480
   ClientLeft      =   3045
   ClientTop       =   3060
   ClientWidth     =   5040
   ControlBox      =   0   'False
   Icon            =   "FRMENVIARBANCO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   240
         Top             =   2640
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         Protocol        =   2
         RemoteHost      =   "ftp.cihtur.com.br"
         RemotePort      =   21
         URL             =   "ftp://cihtur:cihtur0533@ftp.cihtur.com.br"
         UserName        =   "cihtur"
         Password        =   "cihtur0533"
      End
      Begin VB.Timer tmrInternet 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   1440
         Top             =   2760
      End
      Begin VB.TextBox txtResponse 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Voltar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         Picture         =   "FRMENVIARBANCO.frx":A45A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdBaixar 
         Caption         =   "&Enviar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         Picture         =   "FRMENVIARBANCO.frx":A5FC
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEnviarBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBaixar_Click()
  On Error GoTo ErroGeral
  
  'O ARQUIVO A SER ENVIADO DEVERA ESTAR NA RAIZ DO SPEEDY
  'OU ONDE DISCA , SE FOR REDE COMPARTILHADA ENVIAR O ARQUIVO PARA A MAQUINA
  'ONDE ESTA O MODEM
  'O Z FOI MAPEADO E É A MAQUINA ONDE TEM O SPEEDY
  
  
    If Len(Dir("z:\01.jpg")) <> 0 Then
      
      Inet1.Execute "ftp://cihtur:cihtur0533@ftp.cihtur.com.br", "SEND " & _
                    "z:\01.jpg /cihtur/temp/01.jpg"
     
      
    Else
      MsgBox "Não foi possível efetuar a operação!" & vbCrLf & _
             "Contate URGENTE seu fornecedor!", vbCritical, App.ProductName
    
      fraPrincipal.Enabled = True
      Me.MousePointer = vbDefault
    End If

ErroGeral:
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
  Select Case State
    Case 1
      txtResponse = "Procurando o IP especificado..." & Chr(13) & Chr(10) & txtResponse
    
    Case 2
      txtResponse = "IP Encontrado!" & Chr(13) & Chr(10) & txtResponse
    
    Case 3
      txtResponse = "Conectando ao Computador Central..." & Chr(13) & Chr(10) & txtResponse
    
    Case 4
      txtResponse = "Sucesso ao Conectar ao Computador Central!" & Chr(13) & Chr(10) & txtResponse
    
    Case 7
      txtResponse = "Enviando parte dos Dados..." & Chr(13) & Chr(10) & txtResponse
    
    Case 8
      txtResponse = "Parte dos Dados Enviados com Sucesso!" & Chr(13) & Chr(10) & txtResponse
    
    Case 9
      txtResponse = "Desconectando do Computador Central..." & Chr(13) & Chr(10) & txtResponse
    
    Case 10
      txtResponse = "Desconexão Efetuada com Sucesso!" & Chr(13) & Chr(10) & txtResponse
    
    Case 11
      txtResponse = "Erro Ocorrido: " & Inet1.ResponseInfo & Chr(13) & Chr(10) & txtResponse
      fraPrincipal.Enabled = True
      Me.MousePointer = vbDefault
      
    Case 12
      Dim vtData As Variant ' Data variable.
      Dim strData As String: strData = ""
      Dim bDone As Boolean: bDone = False

      ' Get first chunk.
      vtData = Inet1.GetChunk(1024, icString)
      DoEvents
      Do While Not bDone
         strData = strData & vtData
         DoEvents
         ' Get next chunk.
         vtData = Inet1.GetChunk(1024, icString)
         
         If Len(vtData) = 0 Then bDone = True
      Loop
      tmrInternet.Enabled = True
  End Select
End Sub

Private Sub tmrInternet_Timer()
  txtResponse = "Operação Efetuada com Sucesso!" & Chr(13) & Chr(10) & txtResponse
  fraPrincipal.Enabled = True
  Me.MousePointer = vbDefault
  tmrInternet.Enabled = False
End Sub

Private Sub txtResponse_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub
