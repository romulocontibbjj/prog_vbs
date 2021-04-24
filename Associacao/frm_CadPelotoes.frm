VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_CadPelotoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CADASTRO DE PELOTÕES"
   ClientHeight    =   3465
   ClientLeft      =   5775
   ClientTop       =   2625
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4215
   Begin VB.Frame Frame1 
      Caption         =   "PELOTOES"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin MSDataGridLib.DataGrid grd_Pelotoes 
         Bindings        =   "frm_CadPelotoes.frx":0000
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   8388608
         ForeColor       =   12648447
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
         DataMember      =   "Sel_Pelotoes"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "PELOTAO"
            Caption         =   "PELOTAO"
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
            DataField       =   "CIA"
            Caption         =   "CIA"
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
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Cmd_Excluir 
         Caption         =   "&Excluir"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Cmd_Alterar 
         Caption         =   "&Alterar"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_Inserir 
         Caption         =   "&Inserir"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_cia 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txt_Pelotao 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "CIA:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "NOME:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_CadPelotoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Alterar_Click()
If MsgBox("Deseja Alterar os Dados do Pelotão: " & deb_AEJC.rsSel_Pelotoes.Fields("PELOTAO") & " ?", vbYesNo, "ALTERAÇÃO - PELOTÃO") = vbYes Then
    
    deb_AEJC.Up_Pelotoes Trim$(txt_Pelotao.Text), Trim$(txt_cia.Text), deb_AEJC.rsSel_Pelotoes.Fields("PELOTAO")
    
    If deb_AEJC.rsSel_Pelotoes.State = 1 Then deb_AEJC.rsSel_Pelotoes.Close
    deb_AEJC.Sel_Pelotoes
    
    If deb_AEJC.rsSel_Pelotoes.RecordCount > 0 Then
    
        grd_Pelotoes.DataMember = "sel_pelotoes"
        grd_Pelotoes.Refresh
    End If
    
    txt_cia.Text = Empty
    txt_Pelotao = Empty

End If


End Sub

Private Sub Cmd_Excluir_Click()
Dim xpel As String
xpel = deb_AEJC.rsSel_Pelotoes.Fields("pelotao")

deb_AEJC.Del_Pelotoes Trim$(xpel)

MsgBox xpel & " Excluído com Sucesso", vbInformation, "EXCLUSÃO - PELOTÃO"


If deb_AEJC.rsSel_Pelotoes.State = 1 Then deb_AEJC.rsSel_Pelotoes.Close
            deb_AEJC.Sel_Pelotoes
    
            If deb_AEJC.rsSel_Pelotoes.RecordCount > 0 Then
        
                grd_Pelotoes.DataMember = "sel_pelotoes"
                grd_Pelotoes.Refresh
            
            txt_cia.Text = Empty
            txt_Pelotao = Empty
            End If
            
End Sub

Private Sub cmd_Inserir_Click()


If deb_AEJC.rsPesq_Pelotao.State = 1 Then deb_AEJC.rsPesq_Pelotao.Close
    deb_AEJC.Pesq_Pelotao Trim$(txt_Pelotao.Text)

    
    If deb_AEJC.rsPesq_Pelotao.RecordCount = 0 Then
    
        deb_AEJC.In_pelotoes Trim$(txt_Pelotao.Text), Trim$(txt_cia.Text)

        If deb_AEJC.rsSel_Pelotoes.State = 1 Then deb_AEJC.rsSel_Pelotoes.Close
            deb_AEJC.Sel_Pelotoes
    
            If deb_AEJC.rsSel_Pelotoes.RecordCount > 0 Then
        
                grd_Pelotoes.DataMember = "sel_pelotoes"
                grd_Pelotoes.Refresh
            
            MsgBox "Pelotão: " & txt_Pelotao.Text & Chr$(13) & "Cadastrado com Sucesso", vbInformation, "INCLUSÃO - PELOTÃO"
            
    
            txt_cia.Text = Empty
            txt_Pelotao = Empty
        
        End If
    Else
        MsgBox deb_AEJC.rsPesq_Pelotao.Fields("PELOTAO") & " Já Cadastrado no Sistema." & Chr$(13) & "Com a Companhia: " & deb_AEJC.rsPesq_Pelotao.Fields("CIA"), vbCritical, "INCLUSÃO DE PELOTÃO"
    
    End If
    

End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))

End Sub

Private Sub grd_Pelotoes_Click()
txt_Pelotao.Text = deb_AEJC.rsSel_Pelotoes.Fields("PELOTAO")
txt_cia.Text = deb_AEJC.rsSel_Pelotoes.Fields("CIA")

End Sub
