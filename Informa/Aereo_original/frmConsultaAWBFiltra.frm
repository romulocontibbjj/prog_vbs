VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaAWBFiltra 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   1875
   ClientTop       =   4275
   ClientWidth     =   11595
   Icon            =   "frmConsultaAWBFiltra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmConsultaAWBFiltra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlexGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    
    
    
    
If KeyAscii = 13 Then
    xCodAwb = FlexGrid.TextMatrix(FlexGrid.Row, 1)
    
        If de_informa.rsSelAWB.State = 1 Then de_informa.rsSelAWB.Close
        de_informa.SelAWB xCodAwb
        
        If de_informa.rsSelAWB.RecordCount > 0 Then
        
        With frmConsultaAWB
        .TxtFilial.Text = de_informa.rsSelAWB.Fields("filial")
        .TxtSiglaCiaAerea.Text = de_informa.rsSelAWB.Fields("cia")
        .TxtAWB.Caption = de_informa.rsSelAWB.Fields("awb") & "-" & de_informa.rsSelAWB.Fields("dig")
        .TxtSiglaExpedidor.Text = de_informa.rsSelAWB.Fields("siglaorigem")
        .TxtSiglaVIA.Text = de_informa.rsSelAWB.Fields("siglavia")
        .TxtSiglaDestinatario.Text = de_informa.rsSelAWB.Fields("siglades")
        
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadeorigem")) Then .TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadeorigem")) & " - " & de_informa.rsSelAWB.Fields("uforigem") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportoorigem")) & ")"
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadevia")) Then .TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadevia")) & " - " & de_informa.rsSelAWB.Fields("ufvia") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportovia")) & ")"
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadedestino")) Then .TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadedestino")) & " - " & de_informa.rsSelAWB.Fields("ufdestino") & " (" & PriMaiuscula(de_informa.rsSelAWB.Fields("aeroportodestino")) & ")"
        
        .TxtEspecie.Text = de_informa.rsSelAWB.Fields("especie")
        .TxtDescrIATA.Text = de_informa.rsSelAWB.Fields("descrprodsis")
        
            If de_informa.rsSelAWB.Fields("perecivel") = "S" Then
            .TxtPerecivel.Text = "S"
            Else
            .TxtPerecivel.Text = "N"
            End If
            
            If de_informa.rsSelAWB.Fields("retira") = "S" Then
            .TxtClienteRetira.Text = "S"
            Else
            .TxtClienteRetira.Text = "N"
            End If
            
        .TxtModal.Text = de_informa.rsSelAWB.Fields("modal")
        .TxtTipoTaxa.Text = de_informa.rsSelAWB.Fields("tipotaxa")
        .TxtAliquota.Text = de_informa.rsSelAWB.Fields("aliquota")
        .TxtICMS.Text = de_informa.rsSelAWB.Fields("icms")
        .TxtFreteNacional.Text = de_informa.rsSelAWB.Fields("fretenacional")
        .TxtKiloCob.Text = de_informa.rsSelAWB.Fields("kilo")
        .TxtADValorem.Text = de_informa.rsSelAWB.Fields("advalorem")
        .TxtTipoADVAL.Text = de_informa.rsSelAWB.Fields("tipoadval")
        .TxtTXDestino.Text = de_informa.rsSelAWB.Fields("txdestino")
        .TxtTXRedesp.Text = de_informa.rsSelAWB.Fields("txredesp")
        .TxtDescrOutros1.Text = de_informa.rsSelAWB.Fields("descrtxoutros1")
        .TxtOutros1.Text = de_informa.rsSelAWB.Fields("txoutros1")
        .TxtDescrOutros2.Text = de_informa.rsSelAWB.Fields("descrtxoutros2")
        .TxtOutros2.Text = de_informa.rsSelAWB.Fields("txoutros2")
        .TxtFreteTotal.Text = de_informa.rsSelAWB.Fields("fretetotal")
        .TxtVolumes.Text = de_informa.rsSelAWB.Fields("volumes")
        .TxtPesoCubado.Text = de_informa.rsSelAWB.Fields("pesocubado")
        .TxtPesoReal.Text = de_informa.rsSelAWB.Fields("pesoreal")
        .TxtEmissor.Text = de_informa.rsSelAWB.Fields("emissor")
        .TxtEmissao.Text = de_informa.rsSelAWB.Fields("data")
        .TxtHora.Text = de_informa.rsSelAWB.Fields("hora")
        
            If de_informa.rsSelAWB.Fields("cancelado") = "X" Then
            .TxtStatus.Text = "AWB Cancelado"
            Else
            .TxtStatus.Text = ""
            End If
        
        .TxtOBSEmissao.Text = de_informa.rsSelAWB.Fields("obsemissor")
        
        
        .FlexGridNFs.Clear
        .FlexGridNFs.Rows = de_informa.rsSelAWB.RecordCount + 1
        .FlexGridNFs.Cols = 6
        .FlexGridNFs.FixedCols = 0
        .FlexGridNFs.FixedRows = 1
       
        .FlexGridNFs.TextMatrix(0, 0) = "NF"
        .FlexGridNFs.TextMatrix(0, 1) = "Série"
        .FlexGridNFs.TextMatrix(0, 2) = "Valor"
        .FlexGridNFs.TextMatrix(0, 3) = "FilialCTC"
        .FlexGridNFs.TextMatrix(0, 4) = "Remetente"
        .FlexGridNFs.TextMatrix(0, 5) = "Destinatário"
       
        .FlexGridNFs.ColWidth(0) = 700
        .FlexGridNFs.ColWidth(1) = 500
        .FlexGridNFs.ColWidth(2) = 1300
        .FlexGridNFs.ColWidth(3) = 1200
        .FlexGridNFs.ColWidth(4) = 3500
        .FlexGridNFs.ColWidth(5) = 3500
        
        xCodAwb = de_informa.rsSelAWB.Fields("codawb")
        
        
        X = 0
        
            Do Until de_informa.rsSelAWB.EOF
            X = X + 1
            
            If Not IsNull(de_informa.rsSelAWB.Fields("nota")) Then .FlexGridNFs.TextMatrix(X, 0) = de_informa.rsSelAWB.Fields("nota")
            If Not IsNull(de_informa.rsSelAWB.Fields("SERIE")) Then .FlexGridNFs.TextMatrix(X, 1) = de_informa.rsSelAWB.Fields("serie")
            If Not IsNull(de_informa.rsSelAWB.Fields("VALOR")) Then .FlexGridNFs.TextMatrix(X, 2) = Format(de_informa.rsSelAWB.Fields("VALOR"), "##,##0.00")
            If Not IsNull(de_informa.rsSelAWB.Fields("FILIALCTC")) Then .FlexGridNFs.TextMatrix(X, 3) = de_informa.rsSelAWB.Fields("FILIALCTC")
            If Not IsNull(de_informa.rsSelAWB.Fields("REMET_NOME")) Then .FlexGridNFs.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelAWB.Fields("REMET_NOME"))
            If Not IsNull(de_informa.rsSelAWB.Fields("DEST_NOME")) Then .FlexGridNFs.TextMatrix(X, 5) = PriMaiuscula(de_informa.rsSelAWB.Fields("DEST_NOME"))
            
            de_informa.rsSelAWB.MoveNext
            Loop
            
        If de_informa.rsConsultaAWBVolume.State = 1 Then de_informa.rsConsultaAWBVolume.Close
        de_informa.ConsultaAWBVolume xCodAwb
            If de_informa.rsConsultaAWBVolume.RecordCount > 0 Then
            .FlexGridVolumes.Clear
            .FlexGridVolumes.Rows = de_informa.rsConsultaAWBVolume.RecordCount + 1
            .FlexGridVolumes.Cols = 4
            .FlexGridVolumes.FixedCols = 0
            .FlexGridVolumes.FixedRows = 1
            
            .FlexGridVolumes.TextMatrix(0, 0) = "Qtde."
            .FlexGridVolumes.TextMatrix(0, 1) = "Comprimento"
            .FlexGridVolumes.TextMatrix(0, 2) = "Largura"
            .FlexGridVolumes.TextMatrix(0, 3) = "Altura"
            
            .FlexGridVolumes.ColWidth(0) = 500
            .FlexGridVolumes.ColWidth(1) = 1500
            .FlexGridVolumes.ColWidth(2) = 1500
            .FlexGridVolumes.ColWidth(3) = 1500
            
            X = 0
            
                Do Until de_informa.rsConsultaAWBVolume.EOF
                X = X + 1
                .FlexGridVolumes.TextMatrix(X, 0) = de_informa.rsConsultaAWBVolume.Fields("volumes")
                .FlexGridVolumes.TextMatrix(X, 1) = de_informa.rsConsultaAWBVolume.Fields("comprimento")
                .FlexGridVolumes.TextMatrix(X, 2) = de_informa.rsConsultaAWBVolume.Fields("largura")
                .FlexGridVolumes.TextMatrix(X, 3) = de_informa.rsConsultaAWBVolume.Fields("altura")
                
                de_informa.rsConsultaAWBVolume.MoveNext
                Loop
            End If
            
        de_informa.rsSelAWB.MoveFirst
            
        .TxtTotalVM.Text = de_informa.rsSelAWB.Fields("valmerc")
        
        X = 1
        
        If Not IsNull(de_informa.rsSelAWB.Fields("nomeexp")) Then .TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("nomeexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("endexp")) Then .TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("endexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("bairroexp")) Then .TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("bairroexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadexp")) Then .TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("ufexp")) Then .TxtUF(X).Text = de_informa.rsSelAWB.Fields("ufexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("telexp")) Then .TxtTel(X).Text = de_informa.rsSelAWB.Fields("telexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("faxexp")) Then .TxtFAX(X).Text = de_informa.rsSelAWB.Fields("faxexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("cnpjexp")) Then .TxtCGC(X).Text = de_informa.rsSelAWB.Fields("cnpjexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("inscrestexp")) Then .TxtInscrEst(X).Text = de_informa.rsSelAWB.Fields("inscrestexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("cepexp")) Then .TxtCEP(X).Text = de_informa.rsSelAWB.Fields("cepexp")
        If Not IsNull(de_informa.rsSelAWB.Fields("segexp")) Then .TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("segexp"))
        If Not IsNull(de_informa.rsSelAWB.Fields("apoliceexp")) Then .TxtApolice(X).Text = de_informa.rsSelAWB.Fields("apoliceexp")
        
        
        X = 0
        
        If Not IsNull(de_informa.rsSelAWB.Fields("nomedes")) Then .TxtNome(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("nomedes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("enddes")) Then .TxtEnd(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("enddes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("bairrodes")) Then .TxtBairro(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("bairrodes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("cidadedes")) Then .TxtCidade(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("cidadedes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("ufdes")) Then .TxtUF(X).Text = de_informa.rsSelAWB.Fields("ufdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("teldes")) Then .TxtTel(X).Text = de_informa.rsSelAWB.Fields("teldes")
        If Not IsNull(de_informa.rsSelAWB.Fields("faxdes")) Then .TxtFAX(X).Text = de_informa.rsSelAWB.Fields("faxdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("cnpjdes")) Then .TxtCGC(X).Text = de_informa.rsSelAWB.Fields("cnpjdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("inscrestdes")) Then .TxtInscrEst(X).Text = de_informa.rsSelAWB.Fields("inscrestdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("cepdes")) Then .TxtCEP(X).Text = de_informa.rsSelAWB.Fields("cepdes")
        If Not IsNull(de_informa.rsSelAWB.Fields("segdes")) Then .TxtSeguradora(X).Text = PriMaiuscula(de_informa.rsSelAWB.Fields("segdes"))
        If Not IsNull(de_informa.rsSelAWB.Fields("apolicedes")) Then .TxtApolice(X).Text = de_informa.rsSelAWB.Fields("apolicedes")
        End With
        End If
Unload Me
End If
End Sub

Private Sub Form_Load()

Dim xRS As Recordset

        If de_informa.rsSelAWB_NF.State = 1 Then
        Set xRS = de_informa.rsSelAWB_NF
        ElseIf de_informa.rsSelAWB_CTC.State = 1 Then
        Set xRS = de_informa.rsSelAWB_CTC
        End If
        
        FlexGrid.Clear
        
        FlexGrid.Rows = xRS.RecordCount + 1
        FlexGrid.Cols = 7
        FlexGrid.FixedRows = 1
        FlexGrid.FixedCols = 0
        
        FlexGrid.TextMatrix(0, 0) = "NF"
        FlexGrid.TextMatrix(0, 1) = "AWB"
        FlexGrid.TextMatrix(0, 2) = "Data"
        FlexGrid.TextMatrix(0, 3) = "Hora"
        FlexGrid.TextMatrix(0, 4) = "FilialCTC"
        FlexGrid.TextMatrix(0, 5) = "Remetente"
        FlexGrid.TextMatrix(0, 6) = "Destinatário"
        
        FlexGrid.ColWidth(0) = 1500
        FlexGrid.ColWidth(1) = 1300
        FlexGrid.ColWidth(2) = 1000
        FlexGrid.ColWidth(3) = 1000
        FlexGrid.ColWidth(4) = 1000
        FlexGrid.ColWidth(5) = 3000
        FlexGrid.ColWidth(6) = 3000
        
                
        X = 0
        Do Until xRS.EOF
        X = X + 1
        If Not IsNull(xRS.Fields("nota")) Then FlexGrid.TextMatrix(X, 0) = xRS.Fields("nota")
        If Not IsNull(xRS.Fields("codawb")) Then FlexGrid.TextMatrix(X, 1) = xRS.Fields("codawb")
        If Not IsNull(xRS.Fields("data")) Then FlexGrid.TextMatrix(X, 2) = xRS.Fields("data")
        If Not IsNull(xRS.Fields("hora")) Then FlexGrid.TextMatrix(X, 3) = xRS.Fields("Hora")
        If Not IsNull(xRS.Fields("filialctc")) Then FlexGrid.TextMatrix(X, 4) = xRS.Fields("FilialCTC")
        If Not IsNull(xRS.Fields("remet_nome")) Then FlexGrid.TextMatrix(X, 5) = PriMaiuscula(xRS.Fields("remet_nome"))
        If Not IsNull(xRS.Fields("dest_nome")) Then FlexGrid.TextMatrix(X, 6) = PriMaiuscula(xRS.Fields("dest_nome"))
        xRS.MoveNext
        Loop

End Sub
