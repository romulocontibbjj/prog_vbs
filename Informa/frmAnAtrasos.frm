VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAnAtrasos 
   Caption         =   "Analise de Atrasos"
   ClientHeight    =   8235
   ClientLeft      =   3450
   ClientTop       =   1530
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   Begin VB.Frame Frame15 
      Caption         =   "CTCs em Atraso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   600
      TabIndex        =   15
      Top             =   720
      Width           =   11535
      Begin MSDataGridLib.DataGrid gridAtrasos 
         Bindings        =   "frmAnAtrasos.frx":0000
         Height          =   2055
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3625
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
         DataMember      =   "Sel_CtcsAtrasos"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "filialctc"
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
            DataField       =   "emissao"
            Caption         =   "emissao"
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
            DataField       =   "entrega"
            Caption         =   "entrega"
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
            DataField       =   "prz_meta"
            Caption         =   "prz_meta"
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
            DataField       =   "prz_real"
            Caption         =   "prz_real"
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
            DataField       =   "remet_nome"
            Caption         =   "remet_nome"
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
            DataField       =   "cidade_dest"
            Caption         =   "cidade_dest"
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
            DataField       =   "uf_dest"
            Caption         =   "uf_dest"
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
         BeginProperty Column08 
            DataField       =   "dest_nome"
            Caption         =   "dest_nome"
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
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2835,213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2789,858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   3435,024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame24 
      Caption         =   "Ocorrências"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   600
      TabIndex        =   12
      Top             =   3120
      Width           =   8535
      Begin MSDataGridLib.DataGrid GridConsOcorr 
         Bindings        =   "frmAnAtrasos.frx":0019
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8388608
         ForeColor       =   8454143
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "data"
            Caption         =   "data"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
            DataField       =   "cod_ocorr"
            Caption         =   "cd"
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
            DataField       =   "descr_ocorr"
            Caption         =   "ocorrência / descrição"
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
            DataField       =   "usu_ocorr"
            Caption         =   "usuário"
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
            DataField       =   "usu_dataocorr"
            Caption         =   "data inclusão"
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
            DataField       =   "obs_ocorr"
            Caption         =   "obs_ocorr"
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   480,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   269,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3509,858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   14,74
            EndProperty
         EndProperty
      End
      Begin VB.Label lblObs_Ocorr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   720
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   8295
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Ação / Abono do Atraso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   9240
      TabIndex        =   0
      Top             =   3120
      Width           =   2895
      Begin VB.Frame Frame17 
         Caption         =   "Abonar Atraso (Qtde. Dias)"
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
         Top             =   1440
         Width           =   2655
         Begin VB.CommandButton cmdMenos 
            Caption         =   "-"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdMais 
            Caption         =   "+"
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox lblDiasAbono 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Text            =   "00"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdAbonar 
            Caption         =   "Abonar Atraso"
            Height          =   495
            Left            =   1200
            TabIndex        =   2
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Emissão do CTC:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Entrega do CTC:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblEntrega 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Meta/Real/Atraso:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label lblAtrasos 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAnAtrasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
