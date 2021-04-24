VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmduplicados 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1200
   ClientTop       =   1440
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   12780
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid griddupl 
      Bindings        =   "frmduplicados.frx":0000
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8070
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
      DataMember      =   "Sel_BuscaDupliTeste"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "tem_ocorr"
         Caption         =   "st"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "numnf"
         Caption         =   "numnf"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "valmerc"
         Caption         =   "valmerc"
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
         DataField       =   "fretetotal"
         Caption         =   "fretetotal"
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
         DataField       =   "nfs"
         Caption         =   "nfs"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   285,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   480,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2294,929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2610,142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2729,764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   299,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   9180,284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CONSULTA SAC"
      Height          =   495
      Left            =   10440
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Não Cancelar"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar CTC/CTR"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar Processo"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblaguarde 
      AutoSize        =   -1  'True
      Caption         =   "AGUARDE  ....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   4440
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   11640
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cancelados:"
      Height          =   195
      Left            =   10560
      TabIndex        =   5
      Top             =   1440
      Width           =   885
   End
End
Attribute VB_Name = "frmduplicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    lblaguarde.Visible = True
    DoEvents
    
    If de_informa.rsSel_CTRsTeste.State = 1 Then de_informa.rsSel_CTRsTeste.Close
    de_informa.Sel_CTRsTeste
    
    Do Until de_informa.rsSel_CTRsTeste.EOF
    
        If de_informa.rsSel_BuscaDupliTeste.State = 1 Then de_informa.rsSel_BuscaDupliTeste.Close
        de_informa.Sel_BuscaDupliTeste de_informa.rsSel_CTRsTeste.Fields("remet_cgc"), de_informa.rsSel_CTRsTeste.Fields("numnfnum"), de_informa.rsSel_CTRsTeste.Fields("serie")
    
        If de_informa.rsSel_BuscaDupliTeste.RecordCount > 1 Then
            
            Do Until de_informa.rsSel_BuscaDupliTeste.EOF
                If de_informa.rsSel_BuscaDupliTeste.Fields("tem_ocorr") = "N" Or de_informa.rsSel_BuscaDupliTeste.Fields("tem_ocorr") = "2" Then
                    griddupl.DataMember = "Sel_BuscaDupliTeste"
                    griddupl.Refresh
                    lblaguarde.Visible = False
                    Exit Sub
                End If
                de_informa.rsSel_BuscaDupliTeste.MoveNext
            Loop
            
        End If
        
        de_informa.rsSel_CTRsTeste.MoveNext
        
    Loop
    
    lblaguarde.Visible = False
    DoEvents
    
End Sub

Private Sub Command2_Click()
    
    If griddupl.Columns(0) <> "N" Then
        MsgBox "Só é possível Cancelar CTCs com Status N (Sem Posição/Em Trânsito)"
        Exit Sub
    End If
    
    de_informa.Alt_CancCTC xusuario, "CTC EM DUPLICIDADE", griddupl.Columns(3)
    
    tot = Val(tot) + 1
    
    If de_informa.rsSel_BuscaDupliTeste.State = 1 Then de_informa.rsSel_BuscaDupliTeste.Close
    griddupl.DataMember = "Sel_BuscaDupliTeste"
    griddupl.Refresh
    DoEvents
    
    MsgBox "CTC Cancelado !"
    
    If de_informa.rsSel_CTRsTeste.EOF Then
        MsgBox "PROCESSO FINALIZADO !"
    
        lblaguarde.Visible = False
        DoEvents
    
        
        Exit Sub
    End If
    
    de_informa.rsSel_CTRsTeste.MoveNext
    
    
    lblaguarde.Visible = True
    DoEvents
    
    Do Until de_informa.rsSel_CTRsTeste.EOF
    
        If de_informa.rsSel_BuscaDupliTeste.State = 1 Then de_informa.rsSel_BuscaDupliTeste.Close
        de_informa.Sel_BuscaDupliTeste de_informa.rsSel_CTRsTeste.Fields("remet_cgc"), de_informa.rsSel_CTRsTeste.Fields("numnfnum"), de_informa.rsSel_CTRsTeste.Fields("serie")
    
        If de_informa.rsSel_BuscaDupliTeste.RecordCount > 1 Then
            Do Until de_informa.rsSel_BuscaDupliTeste.EOF
                If de_informa.rsSel_BuscaDupliTeste.Fields("tem_ocorr") = "N" Or de_informa.rsSel_BuscaDupliTeste.Fields("tem_ocorr") = "2" Then
                    griddupl.DataMember = "Sel_BuscaDupliTeste"
                    griddupl.Refresh
                    lblaguarde.Visible = False
                    Exit Sub
                End If
                de_informa.rsSel_BuscaDupliTeste.MoveNext
            Loop
        End If
        
        de_informa.rsSel_CTRsTeste.MoveNext
        
        If de_informa.rsSel_CTRsTeste.EOF Then
            MsgBox "PROCESSO FINALIZADO !"
    
            lblaguarde.Visible = False
            DoEvents
    
            
            Exit Sub
        End If
        
    Loop
    
    
    lblaguarde.Visible = False
    DoEvents
    
    
    
    If de_informa.rsSel_CTRsTeste.EOF Then
        MsgBox "PROCESSO FINALIZADO !"
        Exit Sub
    End If
    
End Sub

Private Sub Command3_Click()
    
    
    lblaguarde.Visible = True
    DoEvents
    
    
    If de_informa.rsSel_BuscaDupliTeste.State = 1 Then de_informa.rsSel_BuscaDupliTeste.Close
    griddupl.DataMember = "Sel_BuscaDupliTeste"
    griddupl.Refresh
    DoEvents
    
    If de_informa.rsSel_CTRsTeste.EOF Then
        MsgBox "PROCESSO FINALIZADO !"
    
        lblaguarde.Visible = False
        DoEvents
    
        
        Exit Sub
    End If
    
    de_informa.rsSel_CTRsTeste.MoveNext
    
    Do Until de_informa.rsSel_CTRsTeste.EOF
    
        If de_informa.rsSel_BuscaDupliTeste.State = 1 Then de_informa.rsSel_BuscaDupliTeste.Close
        de_informa.Sel_BuscaDupliTeste de_informa.rsSel_CTRsTeste.Fields("remet_cgc"), de_informa.rsSel_CTRsTeste.Fields("numnfnum"), de_informa.rsSel_CTRsTeste.Fields("serie")
    
        If de_informa.rsSel_BuscaDupliTeste.RecordCount > 1 Then
            Do Until de_informa.rsSel_BuscaDupliTeste.EOF
                If de_informa.rsSel_BuscaDupliTeste.Fields("tem_ocorr") = "N" Or de_informa.rsSel_BuscaDupliTeste.Fields("tem_ocorr") = "2" Then
                    griddupl.DataMember = "Sel_BuscaDupliTeste"
                    griddupl.Refresh
                    lblaguarde.Visible = False
                    Exit Sub
                End If
                de_informa.rsSel_BuscaDupliTeste.MoveNext
            Loop
        End If
        
        de_informa.rsSel_CTRsTeste.MoveNext
        
        If de_informa.rsSel_CTRsTeste.EOF Then
            MsgBox "PROCESSO FINALIZADO !"
    
            lblaguarde.Visible = False
            DoEvents
    
            
            Exit Sub
        End If
        
    Loop
    
    
    lblaguarde.Visible = False
    DoEvents
    
    
    
    If de_informa.rsSel_CTRsTeste.EOF Then
        MsgBox "PROCESSO FINALIZADO !"
        Exit Sub
    End If
    
End Sub

Private Sub Command4_Click()
    frmSac.txtfilial = Mid(griddupl.Columns(3), 1, 2)
    frmSac.txtCTC = Mid(griddupl.Columns(3), 3, 8)
    frmSac.Caption = "SAC - Informação de Transporte - Acompanhamento (chamada)"
    DoEvents
    frmSac.Show
    frmSac.cmbProcurar.SetFocus
    SendKeys "{ENTER}"
End Sub

