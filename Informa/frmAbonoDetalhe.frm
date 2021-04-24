VERSION 5.00
Begin VB.Form frmAbonoDetalhe 
   Caption         =   "Detalhe do Abono / Justificativa do Atraso"
   ClientHeight    =   4230
   ClientLeft      =   2265
   ClientTop       =   1605
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7245
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Em:"
         Height          =   195
         Left            =   5160
         TabIndex        =   23
         Top             =   2280
         Width           =   270
      End
      Begin VB.Label lblDtAbono 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblAbono 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblAbonador 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblReal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblObsAbono 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   6495
      End
      Begin VB.Label lblEntrega 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPrevisao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblMeta 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblDestino 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Motivo / Justificativa:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Abonador:"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   2280
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Abono de Atraso:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
         Height          =   195
         Left            =   4440
         TabIndex        =   9
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Real:"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Entrega em:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Previsão:"
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Meta:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão do CTC:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAbonoDetalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label15_Click()

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbonoDetalhe = Nothing
End Sub
