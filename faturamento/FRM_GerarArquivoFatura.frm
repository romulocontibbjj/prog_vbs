VERSION 5.00
Begin VB.Form FRM_GerarArquivoFatura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerar Arquivo de Faturas - BONAGURA"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command3 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Frame Frame3 
         Caption         =   "Gerar Arquivo"
         Height          =   1455
         Left            =   3600
         TabIndex        =   9
         Top             =   360
         Width           =   3015
         Begin VB.CommandButton cmdProcessarArquivo 
            Caption         =   "Processar ....."
            Height          =   735
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Reenvio de Faturas anteriores"
         Height          =   1935
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3255
         Begin VB.TextBox mskDataFim 
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   1
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox mskDataInicio 
            Height          =   285
            Left            =   480
            MaxLength       =   10
            TabIndex        =   0
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdLimparPeriodo 
            Caption         =   "Limpar Período"
            Height          =   495
            Left            =   840
            TabIndex        =   2
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Ex. dd/mm/aaaa"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Data Fim"
            Height          =   255
            Left            =   1800
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Início"
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "FRM_GerarArquivoFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimparPeriodo_Click()
Dim datade As Date
Dim dataate As Date

Dim DD As String
Dim DA As String

datade = Day(mskDataInicio) & "/" & Month(mskDataInicio) & "/" & Year(mskDataInicio)
dataate = Day(mskDataFim) & "/" & Month(mskDataFim) & "/" & Year(mskDataFim)

If mskDataInicio.Text = "" Or IsDate(mskDataInicio.Text) = False Then
    MsgBox ("Data de Início inválida ..."), vbInformation + vbOKOnly
    mskDataInicio.SetFocus
    Exit Sub
ElseIf mskDataFim.Text = "" Or IsDate(mskDataFim.Text) = False Then
    MsgBox ("Data de Fim inválida ..."), vbInformation + vbOKOnly
    mskDataFim.SetFocus
    Exit Sub
ElseIf dataate < datade Then
    MsgBox ("Data de Fim tem que ser maior que Dta Início ..."), vbInformation + vbOKOnly
    mskDataFim.SetFocus
    
    Exit Sub
Else
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rsSel As New ADODB.Recordset
    Dim sqlUpdate, selString As String
    
    cn.ConnectionString = xstrcon2
    cn.ConnectionTimeout = 30
    cn.Open
    
            
    DD = Year(mskDataInicio) & "/" & Month(mskDataInicio) & "/" & Day(mskDataInicio)
    DA = Year(mskDataFim) & "/" & Month(mskDataFim) & "/" & Day(mskDataFim)
            
    selString = "SELECT * FROM tb_fatura "
    selString = selString & " WHERE emissao >= " & "'" & DD & "'"
    selString = selString & " AND emissao <= " & "'" & DA & "'"
    selString = selString & " AND at_edi_contabil <> " & "'" & "'"
    
    Set rsSel = cn.Execute(selString)
    
    Dim i As Integer
    i = 0
    Do Until rsSel.EOF
        i = i + 1
        rsSel.MoveNext
    Loop
    
    rsSel.MoveFirst
        
    If rsSel.EOF = True Then
        MsgBox ("Nenhuma Informação encontrada"), vbInformation + vbOKOnly
    Else
        If MsgBox("Deseja Realmente Efetuar a atualização de " & i & " Registros ...", vbQuestion + vbYesNo) = vbYes Then
            Do Until rsSel.EOF
                sqlUpdate = "UPDATE tb_fatura SET at_edi_contabil = " & "'" & "'"
                sqlUpdate = sqlUpdate & " WHERE filialfatura = " & "'" & rsSel!filialfatura & "'"
                rsSel.MoveNext
                'Limpa os campos at_edi_contabil
                cn.Execute (sqlUpdate)
            Loop
            MsgBox ("Atualização efetuada com sucesso ..."), vbInformation + vbOKOnly
        End If
    End If

End If

End Sub

Private Sub cmdProcessarArquivo_Click()

cmdProcessarArquivo.Caption = "Aguarde ...."

If MsgBox("Será criado o Arquivo de Fatura da BONAGURA, confirma a Criação ...", vbQuestion + vbYesNo) = vbYes Then
    Call GerarArquivoFaturaBONAGURA
End If

cmdProcessarArquivo.Caption = "Processar"

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

