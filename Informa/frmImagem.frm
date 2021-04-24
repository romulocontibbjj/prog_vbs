VERSION 5.00
Begin VB.Form frmImagem 
   Caption         =   "Imagens Digitalizadas"
   ClientHeight    =   8400
   ClientLeft      =   -3210
   ClientTop       =   2265
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "POD Scanner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11775
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   4800
         TabIndex        =   8
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton cmdProximo 
            Caption         =   ">>"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdAnterior 
            Caption         =   "<<"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblNumImg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   720
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   615
         Left            =   10560
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprTela 
         Height          =   615
         Left            =   9600
         Picture         =   "frmImagem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   690
      End
      Begin VB.CommandButton cmdGravaJpg 
         Caption         =   "Gravar Imagem Abaixo Como Arquivo JPG"
         Height          =   615
         Left            =   6960
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblDataHora 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblTotImg 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total de Imagens Gravadas:"
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial-CTC/CTR:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lblCtc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11775
      Begin VB.Image Image1 
         Height          =   6825
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   11520
      End
   End
End
Attribute VB_Name = "frmImagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnterior_Click()
    Dim Img() As Byte, I As Long
    lblNumImg = Val(lblNumImg) - 1
    de_informa.rsSel_imagem.MovePrevious
    cmdProximo.Enabled = True
    If Val(lblNumImg) = 1 Then
        cmdAnterior.Enabled = False
    End If
    Img = de_informa.rsSel_imagem.Fields("imagem")
    lblDataHora = de_informa.rsSel_imagem.Fields("data")
    Open "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG" For Binary As #2
    Put #2, , Img
    Close #2
    On Error Resume Next
    frmImagem.Image1.Picture = LoadPicture("C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG")
    Kill "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG"
End Sub

Private Sub cmdGravaJpg_Click()
    Dim Img() As Byte, I As Long
    Img = de_informa.rsSel_imagem.Fields("imagem")
    Open "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG" For Binary As #2
    Put #2, , Img
    Close #2
    MsgBox "Imagem Gravada no Diretório C:\" & lblCtc & ".JPG"
End Sub

Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdProximo_Click()
    Dim Img() As Byte, I As Long
    lblNumImg = Val(lblNumImg) + 1
    de_informa.rsSel_imagem.MoveNext
    cmdAnterior.Enabled = True
    If Val(lblNumImg) = Val(lblTotImg) Then
        cmdProximo.Enabled = False
    End If
    Img = de_informa.rsSel_imagem.Fields("imagem")
    lblDataHora = de_informa.rsSel_imagem.Fields("data")
    Open "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG" For Binary As #2
    Put #2, , Img
    Close #2
    On Error Resume Next
    frmImagem.Image1.Picture = LoadPicture("C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG")
    Kill "C:\" & de_informa.rsSel_imagem.Fields("filialctc") & ".JPG"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmImagem = Nothing
End Sub

