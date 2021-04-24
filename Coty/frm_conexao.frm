VERSION 5.00
Begin VB.Form frm_conexao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONEXÃO - IP"
   ClientHeight    =   1470
   ClientLeft      =   5955
   ClientTop       =   4575
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3660
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmd_GravaIp 
         Caption         =   "&Gravar IP"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txt_ip 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "frm_conexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_GravaIp_Click()
Dim xlinha As String
Dim xstrcon As String

If txt_ip.Text = Empty Then

    MsgBox "Digite o IP", vbCritical, "CADASTRO DE IP"
    txt_ip.SetFocus
    
Else

    
    Open "C:\Coty.cnx" For Output As #1

        Print #1, "CNX= " & txt_ip
    
    Close #1

    MsgBox "IP GRAVADO COM SUCESSO", vbInformation, txt_ip
    
    
    Unload MDIForm1
    
    Unload Me

End If



End Sub
