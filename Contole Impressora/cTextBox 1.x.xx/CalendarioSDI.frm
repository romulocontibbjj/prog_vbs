VERSION 5.00
Begin VB.Form CalendarioSDI 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2460
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   45
      ScaleHeight     =   1980
      ScaleWidth      =   2610
      TabIndex        =   2
      Top             =   405
      Width           =   2640
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   75
         X2              =   2535
         Y1              =   340
         Y2              =   340
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   75
         X2              =   2520
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   51
         Top             =   90
         Width           =   315
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   50
         Top             =   90
         Width           =   270
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   825
         TabIndex        =   49
         Top             =   90
         Width           =   240
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   1155
         TabIndex        =   48
         Top             =   90
         Width           =   300
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qui"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   1530
         TabIndex        =   47
         Top             =   90
         Width           =   240
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   1875
         TabIndex        =   46
         Top             =   90
         Width           =   270
      End
      Begin VB.Label LblDiaSemana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sáb"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   2220
         TabIndex        =   45
         Top             =   90
         Width           =   270
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   44
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   43
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   42
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   41
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1545
         TabIndex        =   40
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1890
         TabIndex        =   39
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2250
         TabIndex        =   38
         Top             =   450
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   37
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   495
         TabIndex        =   36
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   840
         TabIndex        =   35
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   1200
         TabIndex        =   34
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1545
         TabIndex        =   33
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   1890
         TabIndex        =   32
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   2250
         TabIndex        =   31
         Top             =   705
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   150
         TabIndex        =   30
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   495
         TabIndex        =   29
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   840
         TabIndex        =   28
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   1200
         TabIndex        =   27
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   1545
         TabIndex        =   26
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   1890
         TabIndex        =   25
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   2250
         TabIndex        =   24
         Top             =   945
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   150
         TabIndex        =   23
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   495
         TabIndex        =   22
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   840
         TabIndex        =   21
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   1200
         TabIndex        =   20
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   1545
         TabIndex        =   19
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   1890
         TabIndex        =   18
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   2250
         TabIndex        =   17
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   150
         TabIndex        =   16
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   495
         TabIndex        =   15
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   840
         TabIndex        =   14
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   1200
         TabIndex        =   13
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   1545
         TabIndex        =   12
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   1890
         TabIndex        =   11
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   2250
         TabIndex        =   10
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   2250
         TabIndex        =   9
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   1890
         TabIndex        =   8
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   39
         Left            =   1545
         TabIndex        =   7
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   1200
         TabIndex        =   6
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   840
         TabIndex        =   5
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   495
         TabIndex        =   4
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label LblDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   150
         TabIndex        =   3
         Top             =   1695
         Width           =   180
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   270
         Left            =   780
         Shape           =   2  'Oval
         Top             =   675
         Width           =   345
      End
   End
   Begin VB.ComboBox CboAno 
      Height          =   315
      Left            =   1695
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   960
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      ItemData        =   "CalendarioSDI.frx":0000
      Left            =   45
      List            =   "CalendarioSDI.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   -30
   End
End
Attribute VB_Name = "CalendarioSDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ControleSolicitante As Control

Dim MesAtual    As Byte
Dim I           As Integer
Dim Hoje        As Byte

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long

Private Function JanelaAtiva() As Boolean
If ControleSolicitante.Parent.MDIChild = True Then 'verifica se tem form MDI porque na presenca de form MDI a funcao GetActiveWindow Ignora os forms filho
     If GetActiveWindow() = Me.hwnd Then
        JanelaAtiva = True
     Else
        If IsChild(GetActiveWindow(), ControleSolicitante.Parent.hwnd) = 1 Then  'verifica se o formulario eh filho
            If Screen.ActiveForm.hwnd = ControleSolicitante.Parent.hwnd Then 'compara o Hwnd informado com o atual
                JanelaAtiva = True
            End If
        End If
    End If
Else
    If GetActiveWindow() = Me.hwnd Then 'se nao for form filho so faz a comparacao
        JanelaAtiva = True
    End If
End If

End Function

Private Sub Form_Activate()

DoEvents
CboMes.SetFocus
Timer1.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set CalendarioSDI = Nothing
End Sub

Function Limpa()

For I = 0 To 41
        LblDia(I).Caption = ""
Next I
End Function

Function PesquisaDia()

Dim Dia  As Byte
Dim Ult  As Byte 'indica o ultima dia do mes
Dim Data As Byte
Dim Outr As Integer 'os dias do outro mes

Limpa
Dia = Weekday("01/" & CboMes.ListIndex + 1 & "/" & CboAno.Text) - 1
Outr = Dia
Data = 0

For I = 0 To Dia - 1
        LblDia(I).ForeColor = Calendario_DiasInativos
        LblDia(I).Caption = Day(DateAdd("d", -Outr, "01/" & CboMes.ListIndex + 1 & "/" & CboAno.Text))
        LblDia(I).Tag = CboMes.ListIndex - 1
        Outr = Outr - 1
Next I

Outr = 0

For I = Dia To 41
        
        Data = Data + 1
        If Not IsDate(Data & "/" & CboMes.ListIndex + 1 & "/" & CboAno.Text) Then
            Ult = I
            Exit For
        End If
        
        LblDia(I).ForeColor = Calendario_DiaAtivos
        LblDia(I).Caption = Format(Data, "00")
        
        If Day(Date) = I Then Hoje = I + Dia - 1

Next I

For I = Ult To 41
        Outr = Outr + 1
        LblDia(I).ForeColor = Calendario_DiasInativos  'dias inativos
        LblDia(I).Caption = Format(Day(DateAdd("d", Outr - 1, "01/" & CboMes.ListIndex + 1 & "/" & CboAno.Text)), "00")
        LblDia(I).Tag = CboMes.ListIndex + 1
Next I
End Function

Private Sub CboAno_Click()
On Error Resume Next
PesquisaDia
End Sub

Private Sub CboMes_Click()
On Error Resume Next
PesquisaDia
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
For I = 1950 To 2050
        CboAno.AddItem I
Next I

CboMes.ListIndex = DateDiff("m", CDate("01/01/" & Year(Date)), Date)
CboAno.ListIndex = DateDiff("yyyy", #1/1/1950#, Date)

PesquisaDia

Shape1.Top = LblDia(Hoje).Top - 30
Shape1.Left = LblDia(Hoje).Left - 90

Me.Top = 0
Me.Left = 0

Me.BackColor = Calendario_FormCorFundo
Picture1.BackColor = Calendario_CorFundo
CboAno.BackColor = Calendario_ComboCorFundo
CboMes.BackColor = Calendario_ComboCorFundo
Shape1.BorderColor = Calendario_Selecionado

For I = 0 To 6
        LblDiaSemana(I).ForeColor = Calendario_DiasSemana
Next I

End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub LblDia_Click(Index As Integer)
Dim Dia As Byte

Dia = LblDia(Index).Caption
Shape1.Top = LblDia(Index).Top - 30
Shape1.Left = LblDia(Index).Left - 90

If LblDia(Index).Tag <> "" Then
    Select Case LblDia(Index).Tag
    Case -1
          CboMes.ListIndex = 11
          CboAno.ListIndex = CboAno.ListIndex - 1
    Case 12
          CboMes.ListIndex = 0
          CboAno.ListIndex = CboAno.ListIndex + 1
    Case Else
          If LblDia(Index).Tag <> "" Then CboMes.ListIndex = LblDia(Index).Tag
    End Select
    
    For I = 0 To 41
        If CByte(LblDia(I).Caption) = Dia Then
            Shape1.Top = LblDia(I).Top - 30
            Shape1.Left = LblDia(I).Left - 90
            Exit For
        End If
    Next I
    
End If

End Sub

Private Sub LblDia_DblClick(Index As Integer)
ControleSolicitante.Text = LblDia(Index).Caption & "/" & Format(CboMes.ListIndex + 1, "00") & "/" & CboAno.Text
Unload Me
End Sub

Private Sub Timer1_Timer()
If Screen.ActiveForm.hwnd <> Me.hwnd Then Unload Me
If JanelaAtiva = False Then
    Unload Me
    ControleSolicitante.SetFocus
End If
End Sub
