VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_final 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro"
   ClientHeight    =   4230
   ClientLeft      =   2145
   ClientTop       =   3495
   ClientWidth     =   6675
   Icon            =   "frm_final.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frm_final.frx":0442
   ScaleHeight     =   4230
   ScaleWidth      =   6675
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   255
      Left            =   5640
      MouseIcon       =   "frm_final.frx":074C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txt_eft 
      DataField       =   "HireDate"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5160
      TabIndex        =   27
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "PROXIMO"
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "ANTERIOR"
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ultimo 
      Caption         =   "ÚLTIMO"
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd_primeiro 
      Caption         =   "Primeiro"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\Nwind.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\Nwind.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employees"
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
   Begin VB.TextBox Text11 
      DataField       =   "PostalCode"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txt_nome 
      DataField       =   "FirstName"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txt_sobre 
      DataField       =   "LastName"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txt_cidade 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txt_end 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   1680
      Width           =   5295
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4320
      TabIndex        =   16
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "HomePhone"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "Country"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "Region"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4320
      TabIndex        =   13
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txt_nasc 
      DataField       =   "BirthDate"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txt_reg 
      BackColor       =   &H00C0FFFF&
      DataField       =   "EmployeeID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label12 
      Caption         =   "Admissão:"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Fone:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Nasc.:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Região:"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "País:"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Extensão:"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Sobrenome:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Registro:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu mem_arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnu_novo 
         Caption         =   "Novo"
      End
      Begin VB.Menu mnu_Gravar 
         Caption         =   "Gravar"
      End
      Begin VB.Menu mnu_editar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnu_confirmar 
         Caption         =   "Confirmar"
      End
      Begin VB.Menu mnu_excluir 
         Caption         =   "Excluir"
      End
      Begin VB.Menu mnu_sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnu_relatorio 
      Caption         =   "&Relatório"
      Begin VB.Menu mnu_funcionarios 
         Caption         =   "Funcionário"
         Begin VB.Menu mnu_nome 
            Caption         =   "Nome"
         End
         Begin VB.Menu mnu_cofigo 
            Caption         =   "Codigo"
         End
      End
   End
End
Attribute VB_Name = "frm_final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_anterior_Click()
With Adodc1.Recordset
.MovePrevious
If .BOF = True Then .MoveFirst
End With



End Sub

Private Sub cmd_gravar_Click()

End Sub

Private Sub cmd_primeiro_Click()
Adodc1.Recordset.MoveFirst

End Sub

Private Sub cmd_proximo_Click()
With Adodc1.Recordset

.MoveNext
If .EOF = True Then .MoveLast

End With



End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub cmd_ultimo_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Form_Load()
travar_tela
Me.Caption = "HOJE É " & UCase(Format(Date, "dddd"))

End Sub

Private Sub mem_excel_Click()
Shell "C:PROGRAM FILES\Microsoft Office\Office\EXCEL.EXE"

End Sub

Private Sub mnu_cofigo_Click()
deb_nwind.rsEmployees.Sort = "EMPLOYEEID"
dtr_funcionarios.Show vbModal

End Sub

Private Sub mnu_confirmar_Click()
travar_tela
End Sub

Private Sub mnu_editar_Click()
destravar_tela

End Sub

Private Sub mnu_excluir_Click()
On Error GoTo TRATAEXCLUIR
    Adodc1.Recordset.Delete
    cmd_proximo_Click
TRATAEXCLUIR:
    If Err.Number = -2147467259 Then
        MsgBox Err.Description
        Adodc1.Refresh
    End If
    
    


End Sub

Private Sub mnu_Gravar_Click()
If txt_nome.Text = Empty Or txt_sobre.Text = Empty Then
    MsgBox " Prencha os Campos" + Chr$(13) + " NOME" + Chr$(13) + " SOBRENOME", vbInformation, "Dados"
    Exit Sub
End If
Adodc1.Recordset.Update 'GRAVAR DADOS
Adodc1.Refresh 'ATUALIZA
Adodc1.Recordset.MoveLast
travar_tela



End Sub

Private Sub mnu_nome_Click()
deb_nwind.rsEmployees.Sort = "FIRSTNAME"
dtr_funcionarios.Show vbModal

End Sub

Private Sub mnu_novo_Click()
destravar_tela
Adodc1.Recordset.AddNew
txt_sobre.SetFocus
cmd_sair.Enabled = True


End Sub

Private Sub mnu_sair_Click()
Unload Me

End Sub

