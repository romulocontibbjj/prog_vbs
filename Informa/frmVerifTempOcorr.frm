VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerifTempOcorr 
   Caption         =   "Schedule de Informação Web Informa e Dispara - V.1.0"
   ClientHeight    =   1620
   ClientLeft      =   1740
   ClientTop       =   1965
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   6045
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin VB.Label lblTempo 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MENSAGEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   990
      TabIndex        =   0
      Top             =   120
      Width           =   4860
   End
End
Attribute VB_Name = "frmVerifTempOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    
    Timer1.Interval = 0
    
    Bar1.Value = 0
    
    If CDbl(lblTempo) > 0 Then
        lblTempo = zeros2(CDbl(lblTempo) - 1, 2)
        Timer1.Interval = 1000
        Exit Sub
    End If
    
    lblMensagem.Caption = "VERIFICANDO ..."
    DoEvents

    If de_informa.rsSel_OcorrTemp.State = 1 Then de_informa.rsSel_OcorrTemp.Close
    de_informa.Sel_OcorrTemp
    
    
    If de_informa.rsSel_OcorrTemp.RecordCount > 0 Then
    
        Bar1.Max = de_informa.rsSel_OcorrTemp.RecordCount
        Bar1.Value = 0
        DoEvents
        xcont = 0
    
        lblMensagem.Caption = "PROCESSANDO ..."
        DoEvents
    
        Do Until de_informa.rsSel_OcorrTemp.EOF
        
            xcont = xcont + 1
            Bar1.Value = xcont
            DoEvents
            xatualprazos = "N"
            
            'informação do WEB INFORMA - Informação de Transportes
            If IsNull(de_informa.rsSel_OcorrTemp.Fields("tipolanc")) Then
                
                If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
                de_informa.Sel_Ctc_SAC de_informa.rsSel_OcorrTemp.Fields("filialctc")
                
                If de_informa.rsSel_BuscaOcorrCod.State = 1 Then de_informa.rsSel_BuscaOcorrCod.Close
                de_informa.Sel_BuscaOcorrCod zeros2(de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), 2)
                
                If de_informa.rsSel_Ctc_SAC.RecordCount > 0 And de_informa.rsSel_BuscaOcorrCod.RecordCount > 0 Then
                
                    If zeros2(de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), 2) <> "01" Then 'ocorrência normal
                        
                        If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then 'se status atual é de entrega trata obonos
                            
                            de_informa.cn_informa.BeginTrans
                                
                                de_informa.ins_ocorr4 de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                                      de_informa.rsSel_Ctc_SAC.Fields("data"), _
                                                      de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                                                      de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                                      de_informa.rsSel_BuscaOcorrCod.Fields("descricao"), _
                                                      de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                      de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                      "EDI-WEB", _
                                                      datahora("datahora")
                                                      
                                de_informa.Alt_AtClienteNFBranco de_informa.rsSel_OcorrTemp.Fields("filialctc")
                            
                                'abono automático de atraso na entrega
                                If de_informa.rsSel_OcorrTemp.Fields("cod_ocorr") = "26" Or _
                                   de_informa.rsSel_OcorrTemp.Fields("cod_ocorr") = "85" Then  'abono automático para atraso
                                    If de_informa.rsSel_CTCEntrega.State = 1 Then de_informa.rsSel_CTCEntrega.Close
                                    de_informa.Sel_CTCEntrega de_informa.rsSel_OcorrTemp.Fields("filialctc")
                                    If de_informa.rsSel_CTCEntrega.RecordCount > 0 Then
                                        If de_informa.rsSel_CTCEntrega.Fields("diasuteis") - _
                                        de_informa.rsSel_CTCEntrega.Fields("abonodias") > _
                                        de_informa.rsSel_CTCEntrega.Fields("prazoentr") Then
                                            'está em atraso, lançar abono automático
                                            xabonodias = de_informa.rsSel_CTCEntrega.Fields("diasuteis") - de_informa.rsSel_CTCEntrega.Fields("prazoentr")
                                            de_informa.Alt_AbonoAtraso xabonodias, "AUTOMATIC", datahora("DATAHORA"), "Abono Automático Devido Ocorrência", de_informa.rsSel_OcorrTemp.Fields("filialctc")
                                        End If
                                    End If
                                End If
                                
                                'ATUALIZA STATUS TB_OCORRTEMP
                                de_informa.Alt_At_OcorrTemp "S", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                                
                            de_informa.cn_informa.CommitTrans
                        
                            If Len(Trim$(de_informa.rsSel_OcorrTemp.Fields("obs_ocorr"))) > 0 Then
                            
                                de_informa.alt_obs_ocorr de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("obs_ocorr")
        
                            End If
                            
                        ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "N" Or _
                               de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
                               
                            de_informa.cn_informa.BeginTrans
                            
                                de_informa.ins_ocorr4 de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                                      de_informa.rsSel_Ctc_SAC.Fields("data"), _
                                                      de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                                                      de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                                      de_informa.rsSel_BuscaOcorrCod.Fields("descricao"), _
                                                      de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                      de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                      "EDI-WEB", _
                                                      datahora("datahora")
                                                      
                                de_informa.Alt_AtClienteNFBranco de_informa.rsSel_OcorrTemp.Fields("filialctc")
                            
                                de_informa.alt_temocorr_sn "2", de_informa.rsSel_OcorrTemp.Fields("filialctc")   'atualiza arquivo de CTC com tem_ocorr = 2
                                
                                If de_informa.rsSel_OcorrTemp.Fields("cod_ocorr") = "39" Or _
                                   de_informa.rsSel_OcorrTemp.Fields("cod_ocorr") = "84" Then  'pre-baixa automática por ser CTC/NF Retido para COnferência
                                   
                                    de_informa.ins_ocorr1 de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                                  de_informa.rsSel_Ctc_SAC.Fields("data"), _
                                                  de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                                                  de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                                  de_informa.rsSel_BuscaOcorrCod.Fields("descricao"), _
                                                  de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                  de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                  de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                  de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                  ".", _
                                                  "AUTO-PREBX", _
                                                  datahora("datahora"), _
                                                  "S", datahora("data")
        
                                    de_informa.alt_temocorr_sn "1", de_informa.rsSel_OcorrTemp.Fields("filialctc")
                                    
                                    xatualprazos = "S"
                                    
                                End If
                                
                                'ATUALIZA STATUS TB_OCORRTEMP
                                de_informa.Alt_At_OcorrTemp "S", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                                
                            de_informa.cn_informa.CommitTrans
                        
                            If Len(Trim$(de_informa.rsSel_OcorrTemp.Fields("obs_ocorr"))) > 0 Then
                            
                                de_informa.alt_obs_ocorr de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("obs_ocorr")
        
                            End If
                                
                        Else 'tem_ocorr = 0 ou C
                        
                            'ATUALIZA STATUS TB_OCORRTEMP
                            de_informa.Alt_At_OcorrTemp "E", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                        
                        End If
                            
                    ElseIf zeros2(de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), 2) = "01" Then 'OCORRÊNCIA DE ENTREGA
                    
                        If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "N" Or _
                            de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then 'SOMENTE SE FOR ESTES STATUS
                    
                            de_informa.cn_informa.BeginTrans
                            
                                de_informa.ins_ocorr1 de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                              de_informa.rsSel_Ctc_SAC.Fields("data"), _
                                              de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                                              de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                              de_informa.rsSel_BuscaOcorrCod.Fields("descricao"), _
                                              de_informa.rsSel_OcorrTemp.Fields("data"), _
                                              de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                              de_informa.rsSel_OcorrTemp.Fields("data"), _
                                              de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                              de_informa.rsSel_OcorrTemp.Fields("recebedor"), _
                                              "EDI-WEB", _
                                              datahora("datahora"), _
                                              "S", datahora("data")
                                              
                                de_informa.alt_temocorr_sn "1", de_informa.rsSel_OcorrTemp.Fields("filialctc")
                                
                                de_informa.Alt_AtClienteNFBranco de_informa.rsSel_OcorrTemp.Fields("filialctc")
                                
                                'ATUALIZA STATUS TB_OCORRTEMP
                                de_informa.Alt_At_OcorrTemp "S", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                                
                                xatualprazos = "S"
    
                            de_informa.cn_informa.CommitTrans
                            
                            If Len(Trim$(de_informa.rsSel_OcorrTemp.Fields("obs_ocorr"))) > 0 Then
                            
                                de_informa.alt_obs_ocorr de_informa.rsSel_OcorrTemp.Fields("filialctc"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("cod_ocorr"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("data"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("hora"), _
                                                         de_informa.rsSel_OcorrTemp.Fields("obs_ocorr")
    
                            End If
                            
                        Else
                            
                            'ATUALIZA STATUS TB_OCORRTEMP
                            de_informa.Alt_At_OcorrTemp "E", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                            
                        End If
                        
                    End If
                    
                Else
                
                    'ATUALIZA STATUS TB_OCORRTEMP
                    de_informa.Alt_At_OcorrTemp "E", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                
                End If
            
                If xatualprazos = "S" Then
                    'foi lançado uma ENTREGA ! calcula os prazos de entrega
                    frmAtualPrazos.lblFilialctc = de_informa.rsSel_OcorrTemp.Fields("filialctc")
                    frmAtualPrazos.Show 1
                End If
            
            'Informação do DISPARA
            Else
                'Status de SISTEMA
                If de_informa.rsSel_OcorrTemp.Fields("tipolanc") = "S" Then
                    'Status de Confirmação de Envio e Recebimento pelo DISPARA
                    If de_informa.rsSel_OcorrTemp.Fields("cod_ocorr") = "12" Then
                    '************************************************************
                        'Busca tb_mem para saber status dos FLAGs
                        If de_informa.rsSel_StatusDisparaTB_Mem.State = 1 Then de_informa.rsSel_StatusDisparaTB_Mem.Close
                        de_informa.Sel_StatusDisparaTB_Mem
                            'verifica se o processo XLM/Web está em uso
                            If de_informa.rsSel_StatusDisparaTB_Mem.Fields("flag_dispara_web") = "0" Then
                                'Seta Status do processo em usu = 1
                                de_informa.Alt_ControleDisparaTb_Mem "1"
                                    de_informa.cn_informa.BeginTrans
                                        'Atualiza Status (at_dispara) do CTC Como Recebido pelo DISPARA (FLAG)
                                        de_informa.Alt_DisparaStatusEnvCTC de_informa.rsSel_OcorrTemp.Fields("filialctc")
                                        de_informa.Alt_At_OcorrTemp "S", de_informa.rsSel_OcorrTemp.Fields("id_controle")
                                    de_informa.cn_informa.CommitTrans
                                'Seta Status do processo em Não Uso = 0
                                de_informa.Alt_ControleDisparaTb_Mem "0"
                            End If
                    '***************************************************************8
                    End If
                End If
            End If
            
            de_informa.rsSel_OcorrTemp.MoveNext
        
        Loop
        
        lblMensagem.Caption = "AGUARDANDO..."
        lblTempo = "05"
        Timer1.Interval = 1000
    
    Else
    
        lblMensagem.Caption = "AGUARDANDO..."
        lblTempo = "05"
        Timer1.Interval = 1000
    
    End If

End Sub
