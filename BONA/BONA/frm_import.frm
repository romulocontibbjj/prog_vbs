VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2835
   ClientTop       =   1575
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   9060
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin MSComctlLib.ProgressBar prg_bar 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   4800
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSDataGridLib.DataGrid grd_dados2 
         Bindings        =   "frm_import.frx":0000
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3836
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         DataMember      =   "Sel_dados"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "LOCALIDADE"
            Caption         =   "LOCALIDADE"
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
            DataField       =   "SIGLA"
            Caption         =   "SIGLA"
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
         BeginProperty Column02 
            DataField       =   "TXMINIMA"
            Caption         =   "TXMINIMA"
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
         BeginProperty Column03 
            DataField       =   "PORKILO"
            Caption         =   "PORKILO"
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
         BeginProperty Column04 
            DataField       =   "GEN_TXCOLETAVALOR"
            Caption         =   "GEN_TXCOLETAVALOR"
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
         BeginProperty Column05 
            DataField       =   "GEN_TXCOLETAEXCED"
            Caption         =   "GEN_TXCOLETAEXCED"
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
         BeginProperty Column06 
            DataField       =   "GEN_TXENTREGAVALOR"
            Caption         =   "GEN_TXENTREGAVALOR"
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
         BeginProperty Column07 
            DataField       =   "GEN_TXENTREGAEXCED"
            Caption         =   "GEN_TXENTREGAEXCED"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2039,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2039,811
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Caption         =   "altera"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   4440
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid grd_dados 
         Bindings        =   "frm_import.frx":0017
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         DataMember      =   "Sel_dados"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "LOCALIDADE"
            Caption         =   "LOCALIDADE"
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
            DataField       =   "SIGLA"
            Caption         =   "SIGLA"
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
         BeginProperty Column02 
            DataField       =   "TXMINIMA"
            Caption         =   "TXMINIMA"
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
         BeginProperty Column03 
            DataField       =   "PORKILO"
            Caption         =   "PORKILO"
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
         BeginProperty Column04 
            DataField       =   "GEN_TXCOLETAVALOR"
            Caption         =   "GEN_TXCOLETAVALOR"
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
         BeginProperty Column05 
            DataField       =   "GEN_TXCOLETAEXCED"
            Caption         =   "GEN_TXCOLETAEXCED"
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
         BeginProperty Column06 
            DataField       =   "GEN_TXENTREGAVALOR"
            Caption         =   "GEN_TXENTREGAVALOR"
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
         BeginProperty Column07 
            DataField       =   "GEN_TXENTREGAEXCED"
            Caption         =   "GEN_TXENTREGAEXCED"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2039,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2039,811
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "pegadados"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Total Reg.:"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Excel As Excel.Application
Dim ExcelA1 As Excel.Worksheet
Dim xlinhamax As Integer, xlinhaatual As Integer
Dim xsigla As String
Dim txminima As Currency
Dim xpeso As Currency
Dim xate As Currency
Dim xexced As Currency


Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.Interactive = False
Excel.Workbooks.Open FileName:="c:\TFA200.xls"
Set ExcelA1 = Excel.Worksheets(1)

xlinhaatual = 0

Do While True

    xlinhaatual = xlinhaatual + 1
    xlinhamax = xlinhamax + 1
    
    If Len(Trim$(ExcelA1.Cells(xlinhaatual, 1))) = 0 And Len(Trim$(ExcelA1.Cells(xlinhaatual, 2))) = 0 And xlinhaatual >= 2 Then
    
        Exit Do
    End If
    
    Label2.Caption = xlinhamax - 1
    Me.Refresh
Loop

Label2.Caption = xlinhamax - 1



Excel.Quit
Set Excel = Nothing
Set ExcelA1 = Nothing



End Sub

Private Sub Command2_Click()
Dim Excel As Excel.Application
Dim ExcelA1 As Excel.Worksheet
Dim xlinhamax As Integer, xlinhaatual As Integer
Dim xsigla As String
Dim txminima As Currency
Dim xpeso As Currency
Dim xate As Currency
Dim xexced As Currency


Set Excel = CreateObject("Excel.Application")
Excel.Visible = True
Excel.Interactive = True
Excel.Workbooks.Open FileName:="c:\TFA200.xls"
Set ExcelA1 = Excel.Worksheets(1)

prg_bar.Min = 0
prg_bar.Max = Label2.Caption


For X = 1 To Label2.Caption

prg_bar.Value = X

xsigla = ExcelA1.Cells(X, 2)
txminima = Round(ExcelA1.Cells(X, 3), 2)
xpeso = Round(ExcelA1.Cells(X, 4), 2)
xate = Round(ExcelA1.Cells(X, 5), 2)
xexced = Round(ExcelA1.Cells(X, 6), 2)

deb_fnac.up_altera txminima, xpeso, xate, xexced, xate, xexced, xsigla

If deb_fnac.rsSel_dados.State = 1 Then deb_fnac.rsSel_dados.Close
deb_fnac.Sel_dados
grd_dados2.DataMember = "sel_dados"
grd_dados2.Refresh


Next

Excel.Quit
Set Excel = Nothing
Set ExcelA1 = Nothing

If deb_fnac.rsSel_dados.State = 1 Then deb_fnac.rsSel_dados.Close
deb_fnac.Sel_dados
grd_dados2.DataMember = "sel_dados"
grd_dados2.Refresh




End Sub

Private Sub Form_Load()

If deb_fnac.rsSel_dados.State = 1 Then deb_fnac.rsSel_dados.Close
deb_fnac.Sel_dados
grd_dados.DataMember = "sel_dados"
grd_dados.Refresh



End Sub

