VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PESQUISA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Manifesto"
   ClientHeight    =   8175
   ClientLeft      =   720
   ClientTop       =   2025
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11805
   Begin VB.Frame PESQUISA 
      Caption         =   "PESQUISA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.TextBox txtctc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   285
         Width           =   1095
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "PROCURAR"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid gridCtc 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   8388608
         ForeColor       =   65535
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "sel_ctc"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label txtfilial 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Filial/CTC:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "PESQUISA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcurar_Click()
Dim xfilialmanifesto As String


xfilialctc = txtFilial & Trim$(txtCtc)
If debManifesto.rssel_ctc.State = 1 Then debManifesto.rssel_ctc.Close
    debManifesto.sel_ctc xfilialctc
    
If debManifesto.rssel_manifesto.RecordCount < 1 Then
    MsgBox "Filial/CTC não Localizados", vbInformation, "Não Localizado"
    
Else
    gridMani.DataMember = "sel_ctc"
    gridMani.Refresh
End If
    


End Sub
