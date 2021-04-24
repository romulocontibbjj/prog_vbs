VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   1395
   ClientTop       =   1755
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6585
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc ado_teste 
      Height          =   735
      Left            =   1440
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=master;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=master;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select top 1 * from tbcliente"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xcnn As New ADODB.Connection
Dim xrs As New ADODB.Recordset

Dim x As Integer
Dim xconex As String
Dim xstring As String


xconex = "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=master;Data Source=."

xcnn.ConnectionString = xconex

xcnn.Open xcnn.ConnectionString

MsgBox xcnn.State

xstring = "select * from tbcliente"
'xstring = "insert into tbcliente values ('RRC2', 8, 'M')"
'xstring = "update tbcliente set codigo = (select codigo +1 where nome  = 'RRC') where nome = 'RRC'"

ado_teste.RecordSource = xstring
ado_teste.Refresh

grid.ColWidth(0) = 100

'Set xrs = xcnn.Execute(xstring)


MsgBox ado_teste.Recordset.RecordCount

Set grid.DataSource = ado_teste
grid.Refresh








End Sub

