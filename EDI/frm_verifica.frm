VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_verifica 
   Caption         =   "ENVIO ELETRÔNICO DE EDI´S"
   ClientHeight    =   6990
   ClientLeft      =   1395
   ClientTop       =   2265
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   6750
   Begin VB.Frame Frame3 
      Caption         =   "ARQUIVOS GERDADOS HOJE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   6495
      Begin MSDataGridLib.DataGrid grd_gerados 
         Bindings        =   "frm_verifica.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4048
         _Version        =   393216
         BackColor       =   12648447
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
         DataMember      =   "Sel_Grd_Logs"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "cgc"
            Caption         =   "cgc"
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
            DataField       =   "cliente"
            Caption         =   "cliente"
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
            DataField       =   "tipodoc"
            Caption         =   "tipodoc"
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
            DataField       =   "horario"
            Caption         =   "horario"
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
         BeginProperty Column06 
            DataField       =   "Obs"
            Caption         =   "Obs"
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
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4065
      Left            =   75
      MouseIcon       =   "frm_verifica.frx":0016
      TabIndex        =   0
      Top             =   75
      Width           =   6615
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&Sair"
         Height          =   315
         Left            =   5475
         TabIndex        =   7
         Top             =   3450
         Width           =   1065
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   5925
         Top             =   2775
      End
      Begin VB.Frame Frame2 
         Caption         =   "VERIFICAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   150
         TabIndex        =   2
         Top             =   2625
         Width           =   3765
         Begin VB.TextBox txt_tempo 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1950
            TabIndex        =   5
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton cmd_desativa 
            Caption         =   "&Desativar"
            Height          =   240
            Left            =   225
            TabIndex        =   4
            Top             =   600
            Width           =   1365
         End
         Begin VB.CommandButton cmd_ativa 
            Caption         =   "&Ativar"
            Height          =   240
            Left            =   225
            TabIndex        =   3
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Tempo (S):"
            Height          =   240
            Left            =   1875
            TabIndex        =   6
            Top             =   300
            Width           =   915
         End
      End
      Begin MSDataGridLib.DataGrid grd_verifica 
         Bindings        =   "frm_verifica.frx":0458
         Height          =   2265
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   3995
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
         DataMember      =   "Sel_EdiDia"
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "cgc"
            Caption         =   "cgc"
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
            DataField       =   "edi"
            Caption         =   "edi"
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
            DataField       =   "cliente"
            Caption         =   "cliente"
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
            DataField       =   "email"
            Caption         =   "email"
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
            DataField       =   "assunto"
            Caption         =   "assunto"
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
            DataField       =   "mensagem"
            Caption         =   "mensagem"
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
            DataField       =   "Salvar"
            Caption         =   "Salvar"
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
            DataField       =   "nomearq"
            Caption         =   "nomearq"
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
         BeginProperty Column09 
            DataField       =   "ddmm"
            Caption         =   "ddmm"
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
         BeginProperty Column10 
            DataField       =   "dia"
            Caption         =   "dia"
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
         BeginProperty Column11 
            DataField       =   "horario"
            Caption         =   "horario"
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
         BeginProperty Column12 
            DataField       =   "semana"
            Caption         =   "semana"
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
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_verifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xvideolar As String

Private Sub cmd_ativa_Click()

If txt_tempo.Text <> Empty Then

Timer1.Interval = Int(txt_tempo.Text) * 1000

Timer1.Enabled = True

txt_tempo.Enabled = False
cmd_ativa.Enabled = False
cmd_desativa.Enabled = True

End If


End Sub

Private Sub cmd_desativa_Click()

Timer1.Enabled = False

txt_tempo.Enabled = True
cmd_ativa.Enabled = True
cmd_desativa.Enabled = False

End Sub

Private Sub cmd_Sair_Click()

MDIForm1.Show

Unload Me


End Sub

Private Sub Form_Load()

If deb_edi.rsSel_EdiDia.State = 1 Then deb_edi.rsSel_EdiDia.Close
    deb_edi.Sel_EdiDia
    
    grd_verifica.DataMember = "sel_edidia"
    grd_verifica.Refresh
    
If deb_edi.rsSel_Grd_Logs.State = 1 Then deb_edi.rsSel_Grd_Logs.Close
    deb_edi.Sel_Grd_Logs Date
    
    grd_gerados.DataMember = "Sel_Grd_Logs"
    grd_gerados.Refresh
    


End Sub

Private Sub Timer1_Timer()
Dim xtime As String
Dim xa As String
Dim xnamearq As String
Dim xe As String
Dim xinstrucao As String
Dim xDia As Integer
Dim xhoje As Integer
Dim xlogs As String

xtime = Mid(Time, 1, 5)

xhoje = Int(Day(Date))

With deb_edi.rsSel_EdiDia

If .State = 1 Then .Close
    deb_edi.Sel_EdiDia


deb_edi.rsSel_EdiDia.MoveFirst

Do Until deb_edi.rsSel_EdiDia.EOF

xDia = .Fields("dia")
'MsgBox Weekday(Date)
    If (Mid(.Fields("semana"), Weekday(Date), 1) = 1) Or (xDia = xhoje) Then
    
        If xtime = .Fields("horario") Then
        
            If .Fields("edi") = "OCOREN" Then
                
                
                
                 If .Fields("cgc") <> "04229761" Then
                        
                        
                        
                        If .Fields("ddmm") = 1 Then
                            xnamearq = .Fields("salvar") & "\" & .Fields("nomearq") & String(2 - Len(Day(Date)), "0") & Day(Date) & String(2 - Len(Month(Date)), "0") & Month(Date) & ".txt"
                    
                        Else
                            xnamearq = .Fields("salvar") & "\" & .Fields("nomearq") & ".txt"
                        End If
                        
                       
                        
                
                End If
                
                frm_verifica.Caption = "Gerando OCOREN - " & .Fields("cliente")
                
                If .Fields("cgc") = "04229761" Then
                
                    xa = OCORENARQ(.Fields("cgc"), .Fields("salvar"))
                Else
                
                 xa = OCORENARQ(.Fields("cgc"), xnamearq)
                
                End If
                
                 xlogs = Xlog(.Fields("CGC"), .Fields("CLIENTE"), .Fields("EDI"), Time, Date, xa)
                 
                
               'ENVIA EMAILS
                If xa = "ARQUIVO GERADO COM SUCESSO" Then
                
                If .Fields("cgc") = "04229761" Then
                    xnamearq = xvideolar
                End If
                
                
                    xe = xmail(.Fields("cliente"), .Fields("email"), .Fields("assunto"), .Fields("mensagem"), xnamearq)
                    
                End If
                
                frm_verifica.Caption = "ENVIO ELETRÔNICO DE EDI´S"
                
            ElseIf .Fields("edi") = "CONEMB" Then
            
                If .Fields("ddmm") = 1 Then
                
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq") & String(2 - Len(Day(Date)), "0") & Day(Date) & String(2 - Len(Month(Date)), "0") & Month(Date) & ".txt"
                Else
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq")
                End If
                    
                frm_verifica.Caption = "Gerando CONEMB - " & .Fields("cliente")
                    
                xa = CONEMB1(.Fields("CLIENTE"), .Fields("CGC"), xnamearq, .Fields("periodo"), .Fields("entrega"), .Fields("cancelados"))
                xlogs = Xlog(.Fields("CGC"), .Fields("CLIENTE"), .Fields("EDI"), Time, Date, xa)
                
                If xa = "ARQUIVO GERADO COM SUCESSO" Then
                    
                    xe = xmail(.Fields("CLIENTE"), .Fields("email"), .Fields("Assunto"), .Fields("mensagem"), xnamearq)
                End If
                
                frm_verifica.Caption = "ENVIO ELETRÔNICO DE EDI´S"
            
            ElseIf .Fields("edi") = "DOCCOB" Then
                
                If .Fields("ddmm") = 1 Then
                
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq") & String(2 - Len(Day(Date)), "0") & Day(Date) & String(2 - Len(Month(Date)), "0") & Month(Date) & ".txt"
                Else
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq")
                End If
                
                frm_verifica.Caption = "Gerando DOCCOB - " & .Fields("cliente")
                
                xa = DOCCOB(xnamearq, .Fields("cgc"))
                
                xlogs = Xlog(.Fields("CGC"), .Fields("CLIENTE"), .Fields("EDI"), Time, Date, xa)
                
                
                If xa = "ARQUIVO GERADO COM SUCESSO" Then
                    
                    xe = xmail(.Fields("CLIENTE"), .Fields("email"), .Fields("Assunto"), .Fields("mensagem"), xnamearq)
                
                
                End If
            
            ElseIf .Fields("EDI") = "NFMEDLEY" Then
            
                
                If .Fields("ddmm") = 1 Then
                
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq") & String(2 - Len(Day(Date)), "0") & Day(Date) & String(2 - Len(Month(Date)), "0") & Month(Date) & ".txt"
                Else
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq")
                End If
                
                frm_verifica.Caption = "Gerando NF - " & .Fields("cliente")
                
                xa = ConembMedley(xnamearq)
                
                xlogs = Xlog(.Fields("CGC"), .Fields("CLIENTE"), .Fields("EDI"), Time, Date, xa)
                
                If xa = "ARQUIVO GERADO COM SUCESSO" Then
                    
                    xe = xmail(.Fields("CLIENTE"), .Fields("email"), .Fields("Assunto"), .Fields("mensagem"), xnamearq)
                                
                End If
                
            ElseIf .Fields("EDI") = "CORREIOS" Then
            
                If .Fields("ddmm") = 1 Then
                
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq") & String(2 - Len(Day(Date)), "0") & Day(Date) & String(2 - Len(Month(Date)), "0") & Month(Date) & ".txt"
                Else
                    xnamearq = .Fields("salvar") & "\" & .Fields("nomearq")
                End If
                
                
                
                xa = xCorreios((Date - 10), Date)
                
                xlogs = Xlog(.Fields("CGC"), .Fields("CLIENTE"), .Fields("EDI"), Time, Date, xa)
                
                xnamearq = xvideolar
                
                If xa = "ARQUIVO GERADO COM SUCESSO" Then
                    
                    xe = xmail(.Fields("CLIENTE"), .Fields("email"), .Fields("Assunto"), .Fields("mensagem"), xnamearq)
                    
                End If
                
            ElseIf .Fields("EDI") = "BONAGURA" Then
            
                xa = xbona()
            
                xlogs = Xlog(.Fields("CGC"), .Fields("CLIENTE"), .Fields("EDI"), Time, Date, xa)
                
                If xa = "ARQUIVO GERADO COM SUCESSO" Then
                    
                    xe = xmail(.Fields("CLIENTE"), .Fields("email"), .Fields("Assunto"), .Fields("mensagem"), xvideolar)
                    
                End If
                
            
            End If
            
        
        End If
        
           frm_verifica.Caption = "ENVIO ELETRÔNICO DE EDI´S"
    End If
    
    frm_verifica.Refresh
    
    deb_edi.rsSel_EdiDia.MoveNext


Me.Refresh

Loop

If deb_edi.rsSel_Grd_Logs.State = 1 Then deb_edi.rsSel_Grd_Logs.Close
    deb_edi.Sel_Grd_Logs Date
    
    grd_gerados.DataMember = "Sel_Grd_Logs"
    grd_gerados.Refresh



End With

End Sub
