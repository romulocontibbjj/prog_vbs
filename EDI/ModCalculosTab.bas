Attribute VB_Name = "ModCalculosTab"
Public Sub calc_TR01(xuf As String, xCidade As String, xvalmerc As Currency, xcodigo As String, flag As String)
    Dim xfrete As Currency, xCim As String
    
    'flag = 'SIMULADO'   //   'EMISSAO'

    'UF, Capital / Interior ou Cidade com percentual sobre o valor de mercadoria (tem valor mínimo)
    
    'procura se há valor de frete específico para a cidade
    If de_informa.rsSel_TR01UFCidade.State = 1 Then de_informa.rsSel_TR01UFCidade.Close
    de_informa.Sel_TR01UFCidade xcodigo, xuf, xCidade
    
    If de_informa.rsSel_TR01UFCidade.RecordCount > 0 Then
        
        'cálculo do frete pela UF e Cidade
        xfrete = xvalmerc * (de_informa.rsSel_TR01UFCidade.Fields("tarifaperc") / 100)
        If xfrete < de_informa.rsSel_TR01UFCidade.Fields("fretemin") Then
            xfrete = de_informa.rsSel_TR01UFCidade.Fields("fretemin")
        End If
        If UCase(flag) = "SIMULADO" Then
            frmSimulado.lblFreteSimul = Format(xfrete, "###,##0.00")
        ElseIf UCase(flag) = "SIMULAFRETE" Then
            If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
                frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfrete, "###,##0.00"))
            Else
                frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfrete, "###,##0.00"))
            End If
        Else
            frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfrete, "###,##0.00"))
        End If
    
    Else
    
        'não encontrou por cidade então procuta por UF/Capital ou Interior
        If de_informa.rsSel_CadCidadePorCidadeUF.State = 1 Then de_informa.rsSel_CadCidadePorCidadeUF.Close
        de_informa.Sel_CadCidadePorCidadeUF xCidade, xuf
        
        If de_informa.rsSel_CadCidadePorCidadeUF.RecordCount < 1 Then
            xCim = "INT"  'interior
        Else
            If de_informa.rsSel_CadCidadePorCidadeUF.Fields("cim") = "I" Then
                xCim = "INT"  'interior
            ElseIf de_informa.rsSel_CadCidadePorCidadeUF.Fields("cim") = "C" Then
                xCim = "CAP"  'capital
            Else
                MsgBox "Erro no Cadastro desta Cidade ! Não Está Identificado se é Capital ou Interior. Procure Suporte Técnico."
                Exit Sub
            End If
        End If
            
        If de_informa.rsSel_TR01UFCim.State = 1 Then de_informa.rsSel_TR01UFCim.Close
        de_informa.Sel_TR01UFCim xcodigo, xuf, xCim
        
        If de_informa.rsSel_TR01UFCim.RecordCount < 1 Then
            MsgBox "Não Está Definido na Tabela Deste Cliente, Valor de Frete para este Estado (Capital ou Interior) !"
            Exit Sub
        Else
            'cálculo do frete pela UF e Cidade
            xfrete = xvalmerc * (de_informa.rsSel_TR01UFCim.Fields("tarifaperc") / 100)
            If xfrete < de_informa.rsSel_TR01UFCim.Fields("fretemin") Then
                xfrete = de_informa.rsSel_TR01UFCim.Fields("fretemin")
            End If
            If UCase(flag) = "SIMULADO" Then
                frmSimulado.lblFreteSimul = Format(xfrete, "###,##0.00")
            ElseIf UCase(flag) = "SIMULAFRETE" Then
                If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
                    frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfrete, "###,##0.00"))
                Else
                    frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfrete, "###,##0.00"))
                End If
            Else
                frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfrete, "###,##0.00"))
            End If
        End If
    End If

End Sub
Public Sub calc_TA01(xuf As String, xCidade As String, xpeso As Currency, xvalmerc As Currency, xcodigo As String, xvia As String, flag As String)
    Dim xfrete As Currency, xfretevalor As Currency, xtxcoleta As Currency, xtxentrredesp As Currency

    If de_informa.rsSel_TA01Sigla.State = 1 Then de_informa.rsSel_TA01Sigla.Close
    de_informa.Sel_TA01Sigla xcodigo, xvia
    
    If de_informa.rsSel_TA01Sigla.RecordCount < 1 Then
        MsgBox "Na Tabela Deste Cliente Não Foi Definido Valor de Frete Para Esta Localidade !"
    Else
    
        'calculo do frete peso
        xfrete = de_informa.rsSel_TA01Sigla.Fields("porkilo") * xpeso
        If xfrete < de_informa.rsSel_TA01Sigla.Fields("txminima") Then
            xfrete = de_informa.rsSel_TA01Sigla.Fields("txminima")
        End If
'        frmEmissaoCTCCTR.txtFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
        
        'calculo do frete valor / advalorem
        xfretevalor = xvalmerc * (de_informa.rsSel_TA01Sigla.Fields("gen_advalorem") / 100)
'        frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
        
        'calculo da taxa de coleta
        xtxcoleta = de_informa.rsSel_TA01Sigla.Fields("gen_txcoletavalor")
        
        If xpeso > de_informa.rsSel_TA01Sigla.Fields("gen_txcoletaate") Then
            xtxcoleta = xtxcoleta + ((xpeso - de_informa.rsSel_TA01Sigla.Fields("gen_txcoletaate")) * _
                                de_informa.rsSel_TA01Sigla.Fields("gen_txcoletaexced"))
        End If
'        frmEmissaoCTCCTR.txtTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
        
        'destino na mesma localidade não cobra redespacho, cobra taxa de entrega
        If de_informa.rsSel_CadLocalidadeSigla.State = 1 Then de_informa.rsSel_CadLocalidadeSigla.Close
        de_informa.Sel_CadLocalidadeSigla xvia
        
        If Trim$(xCidade) = Trim$(de_informa.rsSel_CadLocalidadeSigla.Fields("localidade")) And _
           Trim$(xuf) = Trim$(de_informa.rsSel_CadLocalidadeSigla.Fields("uf")) Then  'é a mesma cidade/uf. cobrar taxa de entrega
           
            'calculo da taxa de entrega
            xtxentrredesp = de_informa.rsSel_TA01Sigla.Fields("gen_txentregavalor")
            
            If xpeso > de_informa.rsSel_TA01Sigla.Fields("gen_txentregaate") Then
                xtxentrredesp = xtxentrredesp + ((xpeso - de_informa.rsSel_TA01Sigla.Fields("gen_txentregaate")) * _
                                    de_informa.rsSel_TA01Sigla.Fields("gen_txentregaexced"))
            End If
            
        
        Else   'destino não é na mesma localidade, cobrar redespacho
            
            'calculo da taxa de redespacho
            xtxentrredesp = de_informa.rsSel_TA01Sigla.Fields("txredesp_valor")
            
            If xpeso > de_informa.rsSel_TA01Sigla.Fields("txredesp_ate") Then
                xtxentrredesp = xtxentrredesp + ((xpeso - de_informa.rsSel_TA01Sigla.Fields("txredesp_ate")) * _
                                    de_informa.rsSel_TA01Sigla.Fields("txredesp_exced"))
            End If
        
        End If
        
'        frmEmissaoCTCCTR.txtTxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
        
        If UCase(flag) = "SIMULADO" Then
            frmSimulado.lblFreteSimul = Format(xfrete + xfretevalor + xtxcoleta + xtxentrredesp, "##,###,##0.00")
        ElseIf UCase(flag) = "SIMULAFRETE" Then
            If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
                frmSimulaFrete.txtAFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
                frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
                frmSimulaFrete.txtATxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmSimulaFrete.txtATxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
            Else
                frmSimulaFrete.txtRFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
                frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
                frmSimulaFrete.txtRTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmSimulaFrete.txtRTxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
            End If
        Else
            frmEmissaoCTCCTR.txtFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
            frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
            frmEmissaoCTCCTR.txtTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
            frmEmissaoCTCCTR.txtTxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
        End If
        
    End If
    
End Sub
Public Sub calc_TA02(xpeso As Currency, xvalmerc As Currency, xcodigo As String, flag As String)
    Dim xfrete As Currency, xfretevalor As Currency, xtxcoleta As Currency, xtxentrredesp As Currency
    
    If de_informa.rsSel_TA02PorPeso.State = 1 Then de_informa.rsSel_TA02PorPeso.Close
    de_informa.Sel_TA02PorPeso xcodigo, xpeso, xpeso
    
    If de_informa.rsSel_TA02PorPeso.RecordCount < 1 Then  'não encontrou esta faixa de peso
        MsgBox "Na Tabela de Preço não está Definido Valor de Frete para Este Peso !"
        Exit Sub
    Else
        'calcula o frete peso
        If de_informa.rsSel_TA02PorPeso.Fields("valormin") > 0 Then
            xfrete = de_informa.rsSel_TA02PorPeso.Fields("valormin")
        Else
            'MODO DE CÁLCULO POR KILO EXCEDENTE DA ÚLTIMA FAIXA
            xfrete = (xpeso - (de_informa.rsSel_TA02PorPeso.Fields("pesode") - 0.1)) * _
                      de_informa.rsSel_TA02PorPeso.Fields("porkilo")
            xfrete = xfrete + de_informa.rsSel_TA02PorPeso.Fields("complemento")
        End If
                      
        'calcula o fretevalor - advalorem
        xfretevalor = xvalmerc * (de_informa.rsSel_TA02PorPeso.Fields("gen_advalorem") / 100)
        
        'calcula o taxa de coleta
        xtxcoleta = de_informa.rsSel_TA02PorPeso.Fields("gen_txcoletavalor")
        If xpeso > de_informa.rsSel_TA02PorPeso.Fields("gen_txcoletaate") Then
            xtxcoleta = xtxcoleta + ((xpeso - de_informa.rsSel_TA02PorPeso.Fields("gen_txcoletaate")) * _
                        de_informa.rsSel_TA02PorPeso.Fields("gen_txcoletaexced"))
        End If
        
        'calcula o taxa de entrega
        xtxentrredesp = de_informa.rsSel_TA02PorPeso.Fields("gen_txentregavalor")
        If xpeso > de_informa.rsSel_TA02PorPeso.Fields("gen_txentregaate") Then
            xtxentrredesp = xtxentrredesp + ((xpeso - de_informa.rsSel_TA02PorPeso.Fields("gen_txentregaate")) * _
                        de_informa.rsSel_TA02PorPeso.Fields("gen_txentregaexced"))
        End If
        
        If UCase(flag) = "SIMULADO" Then
            frmSimulado.lblFreteSimul = Format(xfrete + xfretevalor + xtxcoleta + xtxentrredesp, "##,###,##0.00")
        ElseIf UCase(flag) = "SIMULAFRETE" Then
            If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
                frmSimulaFrete.txtAFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
                frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
                frmSimulaFrete.txtATxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmSimulaFrete.txtATxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
            Else
                frmSimulaFrete.txtRFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
                frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
                frmSimulaFrete.txtRTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmSimulaFrete.txtRTxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
            End If
        Else
            frmEmissaoCTCCTR.txtFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
            frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
            frmEmissaoCTCCTR.txtTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
            frmEmissaoCTCCTR.txtTxRedesp = SoNumeros(Format(xtxentrredesp, "###,##0.00"))
        End If
        
    End If

End Sub
Public Sub calc_TG01(xvalmerc As Currency, xcodigo As String, flag As String)
    Dim xfrete As Currency, xfretevalor As Currency

    If de_informa.rsSel_TG01CodigoAtiva.State = 1 Then de_informa.rsSel_TG01CodigoAtiva.Close
    de_informa.Sel_TG01CodigoAtiva xcodigo
    
    xfrete = xvalmerc * (de_informa.rsSel_TG01CodigoAtiva.Fields("fretepeso") / 100)
    If xfrete < de_informa.rsSel_TG01CodigoAtiva.Fields("freteminimo") Then
        xfrete = de_informa.rsSel_TG01CodigoAtiva.Fields("freteminimo")
    End If

    xfretevalor = xvalmerc * (de_informa.rsSel_TG01CodigoAtiva.Fields("fretevalor") / 100)
    
    If UCase(flag) = "SIMULADO" Then
        frmSimulado.lblFreteSimul = Format(xfrete + xfretevalor, "##,###,##0.00")
    ElseIf UCase(flag) = "SIMULAFRETE" Then
        If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
            frmSimulaFrete.txtAFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
            frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
        Else
            frmSimulaFrete.txtRFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
            frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
        End If
    Else
        frmEmissaoCTCCTR.txtFretePeso = SoNumeros(Format(xfrete, "###,##0.00"))
        frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
    End If
    
End Sub
Public Sub calc_TR02(xuf As String, xCidade As String, xvalmerc As Currency, xpeso As Currency, xcodigo As String, flag As String)
    Dim xfretepeso As Currency, xfretevalor As Currency, xtxcoleta As Currency, xtxentregared As Currency

    'UF, Capital / Interior ou Cidade com percentual sobre o valor de mercadoria (tem valor mínimo)
    
    'procura se há valor de frete específico para a cidade
    If de_informa.rsSel_TR02UFCidadePeso.State = 1 Then de_informa.rsSel_TR02UFCidadePeso.Close
    de_informa.Sel_TR02UFCidadePeso xcodigo, xuf, xCidade, xpeso, xpeso
    
    If de_informa.rsSel_TR02UFCidadePeso.RecordCount > 0 Then
        
        'cálculo do frete pela UF e Cidade
        If de_informa.rsSel_TR02UFCidadePeso.Fields("fretepeso") > 0 Then
            xfretepeso = de_informa.rsSel_TR02UFCidadePeso.Fields("fretepeso")
            If xpeso * de_informa.rsSel_TR02UFCidadePeso.Fields("porkilo") > xfretepeso Then
                xfretepeso = xpeso * de_informa.rsSel_TR02UFCidadePeso.Fields("porkilo")
            End If
        Else
            'MODO DE CÁLCULO POR KILO E NÃO EXCEDENTE DA ÚLTIMA FAIXA
            xfretepeso = xpeso * de_informa.rsSel_TR02UFCidadePeso.Fields("porkilo")
            xfretepeso = xfretepeso + de_informa.rsSel_TR02UFCidadePeso.Fields("complemento")
        End If
        
        'calculo da taxa de coleta
        xtxcoleta = de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txcoletavalor")
        
        If xpeso > de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txcoletaate") Then
            xtxcoleta = xtxcoleta + ((xpeso - de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txcoletaate")) * _
                                de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txcoletaexced"))
        End If
        
        'calculo da taxa de entrega
        xtxentregared = de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txentregavalor")
        
        If xpeso > de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txentregaate") Then
            xtxentregared = xtxentregared + ((xpeso - de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txentregaate")) * _
                                de_informa.rsSel_TR02UFCidadePeso.Fields("gen_txentregaexced"))
        End If
        
        'calculo do Frete Valor
        If de_informa.rsSel_TR02UFCidadePeso.Fields("fretevalor") > 0 Then
            xfretevalor = xvalmerc * (de_informa.rsSel_TR02UFCidadePeso.Fields("fretevalor") / 100)
        End If
        
        If UCase(flag) = "SIMULADO" Then
            frmSimulado.lblFreteSimul = Format(xfretepeso + xtxcoleta + xtxentregared + xfretevalor, "##,###,##0.00")
        ElseIf UCase(flag) = "SIMULAFRETE" Then
            If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
                frmSimulaFrete.txtAFretePeso = SoNumeros(Format(xfretepeso, "###,##0.00"))
                frmSimulaFrete.txtATxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmSimulaFrete.txtATxRedesp = SoNumeros(Format(xtxentregared, "###,##0.00"))
                frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
            Else
                frmSimulaFrete.txtRFretePeso = SoNumeros(Format(xfretepeso, "###,##0.00"))
                frmSimulaFrete.txtRTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmSimulaFrete.txtRTxRedesp = SoNumeros(Format(xtxentregared, "###,##0.00"))
                frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
            End If
        Else
            frmEmissaoCTCCTR.txtFretePeso = SoNumeros(Format(xfretepeso, "###,##0.00"))
            frmEmissaoCTCCTR.txtTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
            frmEmissaoCTCCTR.txtTxRedesp = SoNumeros(Format(xtxentregared, "###,##0.00"))
            frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
        End If
        
    Else
    
        'não encontrou por cidade então procuta por UF/Capital ou Interior
        If de_informa.rsSel_CadCidadePorCidadeUF.State = 1 Then de_informa.rsSel_CadCidadePorCidadeUF.Close
        de_informa.Sel_CadCidadePorCidadeUF xCidade, xuf
        
        If de_informa.rsSel_CadCidadePorCidadeUF.RecordCount < 1 Then
            xCim = "INT"  'interior
        Else
            If de_informa.rsSel_CadCidadePorCidadeUF.Fields("cim") = "I" Then
                xCim = "INT"  'interior
            ElseIf de_informa.rsSel_CadCidadePorCidadeUF.Fields("cim") = "C" Then
                xCim = "CAP"  'capital
            Else
                MsgBox "Erro no Cadastro desta Cidade ! Não Está Identificado se é Capital ou Interior. Procure Suporte Técnico."
                Exit Sub
            End If
        End If
            
        If de_informa.rsSel_TR02UFCimPeso.State = 1 Then de_informa.rsSel_TR02UFCimPeso.Close
        de_informa.Sel_TR02UFCimPeso xcodigo, xuf, xCim, xpeso, xpeso
        
        If de_informa.rsSel_TR02UFCimPeso.RecordCount < 1 Then
            MsgBox "Não Está Definido na Tabela Deste Cliente, Valor de Frete para este Estado (Capital ou Interior) !"
            If UCase(flag) = "SIMULADO" Then
                frmSimulado.lblFreteSimul = "0,00"
            End If
            Exit Sub
        Else
            'cálculo do frete pela UF e Cidade
            If de_informa.rsSel_TR02UFCimPeso.Fields("fretepeso") > 0 Then
                xfretepeso = de_informa.rsSel_TR02UFCimPeso.Fields("fretepeso")
                If xpeso * de_informa.rsSel_TR02UFCimPeso.Fields("porkilo") > xfretepeso Then
                    xfretepeso = xpeso * de_informa.rsSel_TR02UFCimPeso.Fields("porkilo")
                End If
            Else
                'MODO DE CÁLCULO POR KILO E NÃO EXCEDENTE DA ÚLTIMA FAIXA
                xfretepeso = xpeso * de_informa.rsSel_TR02UFCimPeso.Fields("porkilo")
                xfretepeso = xfretepeso + de_informa.rsSel_TR02UFCimPeso.Fields("complemento")
            End If
            
            'calculo da taxa de coleta
            xtxcoleta = de_informa.rsSel_TR02UFCimPeso.Fields("gen_txcoletavalor")
            
            If xpeso > de_informa.rsSel_TR02UFCimPeso.Fields("gen_txcoletaate") Then
                xtxcoleta = xtxcoleta + ((xpeso - de_informa.rsSel_TR02UFCimPeso.Fields("gen_txcoletaate")) * _
                                    de_informa.rsSel_TR02UFCimPeso.Fields("gen_txcoletaexced"))
            End If
            
            'calculo da taxa de entrega
            xtxentregared = de_informa.rsSel_TR02UFCimPeso.Fields("gen_txentregavalor")
            
            If xpeso > de_informa.rsSel_TR02UFCimPeso.Fields("gen_txentregaate") Then
                xtxentregared = xtxentregared + ((xpeso - de_informa.rsSel_TR02UFCimPeso.Fields("gen_txentregaate")) * _
                                    de_informa.rsSel_TR02UFCimPeso.Fields("gen_txentregaexced"))
            End If
            
            'calculo do Frete Valor
            If de_informa.rsSel_TR02UFCimPeso.Fields("fretevalor") > 0 Then
                xfretevalor = xvalmerc * (de_informa.rsSel_TR02UFCimPeso.Fields("fretevalor") / 100)
            End If
            
            If UCase(flag) = "SIMULADO" Then
                frmSimulado.lblFreteSimul = Format(xfretepeso + xtxcoleta + xtxentregared + xfretevalor, "##,###,##0.00")
            ElseIf UCase(flag) = "SIMULAFRETE" Then
                If Mid$(frmSimulaFrete.lblModalAtual, 1, 1) = "A" Then
                    frmSimulaFrete.txtAFretePeso = SoNumeros(Format(xfretepeso, "###,##0.00"))
                    frmSimulaFrete.txtATxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                    frmSimulaFrete.txtATxRedesp = SoNumeros(Format(xtxentregared, "###,##0.00"))
                    frmSimulaFrete.txtAFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
                Else
                    frmSimulaFrete.txtRFretePeso = SoNumeros(Format(xfretepeso, "###,##0.00"))
                    frmSimulaFrete.txtRTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                    frmSimulaFrete.txtRTxRedesp = SoNumeros(Format(xtxentregared, "###,##0.00"))
                    frmSimulaFrete.txtRFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
                End If
            Else
                frmEmissaoCTCCTR.txtFretePeso = SoNumeros(Format(xfretepeso, "###,##0.00"))
                frmEmissaoCTCCTR.txtTxColeta = SoNumeros(Format(xtxcoleta, "###,##0.00"))
                frmEmissaoCTCCTR.txtTxRedesp = SoNumeros(Format(xtxentregared, "###,##0.00"))
                frmEmissaoCTCCTR.txtFreteValor = SoNumeros(Format(xfretevalor, "###,##0.00"))
            End If
            
        End If
    End If

End Sub
