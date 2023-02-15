#include "TOTVS.CH"

/*
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦ Programa ¦ XPOSEST  ¦ Autor ¦ Josue Barbosa         ¦ Data ¦ 29/03/22 ¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦Descrição ¦ Posição de estoque customizado para considerar os empenhos ¦¦¦
¦¦¦          ¦ de SA                                                      ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦ Uso      ¦ Local                                                      ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
*/

User Function XPOSEST()

	Local aParam 	:= {}
	Local aRet		:= {}
	Local aEmp      := {}
	Local oProcess

	/* Definicao dos parâmetros */
	aAdd( aParam ,{1, "Período Inicial: ", CTOD("//"), PesqPict("SCP","CP_EMISSAO"),,,, 50, .T.})
	aAdd( aParam ,{1, "Período Final: "  , CTOD("//"), PesqPict("SCP","CP_EMISSAO"),,,, 50, .T.})
	aAdd( aParam, {1, "Produto De"       , Space(250),  "", ".T.", "SB1", ".T.", 100,  .F.})
	aAdd( aParam, {1, "Armazem"          , Space(250),  "", ".T.", "NNR", ".T.", 100,  .F.})
	aAdd( aParam, {1, "Grupo"            , Space(250),  "", ".T.", "SFM", ".T.", 100,  .F.})

	If ParamBox(aParam, "Posição de estoque - Customizado", aRet)


		aEmp := U_xSelEmp(cEmpAnt)

		If Len(aEmp) = 0
			MsgInfo("Favor selecionar pelo menos uma empresa/filial para que seja possivel prosseguir com a extração!","Atenção!")
			Return
		Endif

		oProcess := MsNewProcess():New({|| XPROCEST(MV_PAR01,MV_PAR02,MV_PAR03,MV_PAR04,MV_PAR05, aEmp, oProcess)}, "Processando...", "Realizando Extração...", .T.)
		oProcess:Activate()

        //Processa( {|| XPROCEST(MV_PAR01,MV_PAR02,MV_PAR03,MV_PAR04,MV_PAR05, aEmp, oProcess)}, "Aguarde...",  "Realizando Extração...")

	EndIf

Return


Static Function XPROCEST(_dDataini,_dDataFim,_cProduto,_cArmazem,_cGrupo,_aEmp, _oProcess)

	Local cTabQry      := GetNextAlias()
	Local oState
	Local y            := 0
    Local cCaminho     := "C:\Exportação estoque\"
	
    Private cArq       := ""
	Private cWorkSheet := ""
	Private cTable     := ""
	Private oFwMsEx    := Nil
	Private aRet       := {}
	Private aInfoPlan  := {}

	Public nHErp01
	Public nHandle

	oFwMsEx := FWMsExcelEx():New()

	cWorkSheet := "Status"
	cTable     := "POSIÇÃO DE ESTOQUE - ("+DTOC(Date())+" - "+Time()+") "

	oFwMsEx:AddWorkSheet( cWorkSheet )
	oFwMsEx:AddTable( cWorkSheet, cTable )

	oFwMsEx:AddColumn( cWorkSheet, cTable , "Filial"               , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Codigo"               , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Tipo"                 , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Grupo"                , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Descrição"            , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "U.M."                 , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Armazém"              , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Descrição Armazém"    , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Saldo em Estoque"     , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Empenho - SA"         , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Empenho - OP"         , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Empenho - PV"         , 1,1)
	oFwMsEx:AddColumn( cWorkSheet, cTable , "Estoque Disponível"   , 1,1)

	_oProcess:SetRegua1(Len(_aEmp))

	For y:= 1 To Len(_aEmp)

		_oProcess:IncRegua1("Processando Empresa: "+_aEmp[y][2]+" - "+_aEmp[y][4])

		cQuery := " SELECT FILIAL, "
		cQuery += "        PRODUTO, "
		cQuery += "        TIPO, "
		cQuery += "        GRUPO, "
		cQuery += "        DESCRICAO, "
		cQuery += "        UM, "
		cQuery += "        ARMAZEM, "
		cQuery += "        DESCRICAO_ARM, "
		cQuery += "        SALDO, "
		cQuery += "        (CASE WHEN EMPENHO_SA IS NULL THEN 0 ELSE EMPENHO_SA END) AS EMPENHO_SA, "
		cQuery += "        EMPENHO_OP, "
		cQuery += "        EMPENHO_PV, "
		cQuery += "        (CASE WHEN EMPENHO_SA IS NULL THEN (SALDO - EMPENHO_OP - EMPENHO_PV) ELSE (SALDO - EMPENHO_SA - EMPENHO_OP - EMPENHO_PV) END) AS ESTOQUE_DISP "
		cQuery += " FROM "
		cQuery += "   (SELECT B2_FILIAL AS FILIAL, "
		cQuery += "           B1_COD AS PRODUTO, "
		cQuery += "           B1_TIPO AS TIPO, "
		cQuery += "           B1_GRUPO AS GRUPO, "
		cQuery += "           B1_DESC AS DESCRICAO, "
		cQuery += "           B1_UM AS UM, "
		cQuery += "           B2_LOCAL AS ARMAZEM, "
		cQuery += "           NNR_DESCRI AS DESCRICAO_ARM, "
		cQuery += "           B2_QATU AS SALDO, "
		cQuery += "      (SELECT SUM(CP_QUANT - CP_QUJE) "
		cQuery += "       FROM SCP"+_aEmp[y][1]+"0 SCP "
		cQuery += "       WHERE CP_FILIAL = SB2.B2_FILIAL "
		cQuery += "         AND CP_EMISSAO BETWEEN '"+DTOS(_dDataini)+"' AND '"+DTOS(_dDataFim)+"' "
		cQuery += "         AND CP_PRODUTO = SB2.B2_COD "
		cQuery += "         AND CP_LOCAL   = SB2.B2_LOCAL "
        cQuery += "         AND CP_STATUS  <> 'E' "
        cQuery += "         AND CP_PREREQU = 'S'  "
		cQuery += "         AND SCP.D_E_L_E_T_ <> '*') AS EMPENHO_SA, "
		cQuery += "           B2_QEMP AS EMPENHO_OP, "
		cQuery += "           B2_RESERVA AS EMPENHO_PV, "
		cQuery += "           0 AS ESTOQUE_DISP "
		cQuery += "    FROM SB2"+_aEmp[y][1]+"0 SB2 "
		cQuery += "    LEFT JOIN SB1"+_aEmp[y][1]+"0 SB1 ON (B1_COD = B2_COD "
		cQuery += "                             AND SB1.D_E_L_E_T_ <> '*') "
		cQuery += "    LEFT JOIN NNR"+_aEmp[y][1]+"0 NNR ON (NNR_FILIAL = B2_FILIAL "
		cQuery += "                             AND NNR_CODIGO = B2_LOCAL "
		cQuery += "                             AND NNR.D_E_L_E_T_ <> '*') "
		cQuery += "    WHERE B2_FILIAL = '"+_aEmp[y][2]+"' "
        cQuery += "      AND B2_LOCAL NOT IN (' ','07') "

		If !Empty(_cProduto)
            cQuery += ConvCpo("B2_COD",Alltrim(_cProduto))       
		Endif

		If !Empty(_cArmazem)
            cQuery += ConvCpo("B2_LOCAL",Alltrim(_cArmazem))       
		Endif

		If !Empty(_cGrupo)
            cQuery += ConvCpo("B1_GRUPO",Alltrim(_cGrupo))       
		Endif

        cQuery += "      AND SB2.D_E_L_E_T_ <> '*' ) TRB "
		cQuery += " ORDER BY FILIAL, "
		cQuery += "          PRODUTO, "
		cQuery += "          ARMAZEM "

		oState := FWPreparedStatement():New()
		oState:SetQuery(cQuery)

		MPSysOpenQuery(cQuery,cTabQry)
		oState:Destroy()

		nTot := Contar(cTabQry,"!Eof()")

		(cTabQry)->(DbGoTop())

		_oProcess:SetRegua2(nTot)
        //ProcRegua(nTot)

		While (cTabQry)->(!Eof())

			_oProcess:IncRegua2("Processando produto -> "+Alltrim((cTabQry)->PRODUTO)+" - "+Alltrim((cTabQry)->DESCRICAO))
            //IncProc("Processando produto -> "+Alltrim((cTabQry)->PRODUTO)+" - "+Alltrim((cTabQry)->DESCRICAO))

			oFwMsEx:AddRow( cWorkSheet, cTable,{(cTabQry)->FILIAL+" - "+Alltrim(_aEmp[y][4]),; // Filial
			(cTabQry)->PRODUTO,;                  // Codigo Produto
			(cTabQry)->TIPO,;                     // Tipo Produto
			(cTabQry)->GRUPO,;                    // Grupo Produto
			Alltrim((cTabQry)->DESCRICAO),;       // Descrição
			(cTabQry)->UM,;                       // Unidade de Medida
			(cTabQry)->ARMAZEM,;                  // Armazem
			Alltrim((cTabQry)->DESCRICAO_ARM),;   // Descrição Armazém
			cValToChar((cTabQry)->SALDO),;        // Saldo
			cValToChar((cTabQry)->EMPENHO_SA),;   // Empenho SA
			cValToChar((cTabQry)->EMPENHO_OP),;   // Empenho OP
			cValToChar((cTabQry)->EMPENHO_PV),;   // Empenho PV
			cValToChar((cTabQry)->ESTOQUE_DISP)}) // Estoque disponível

			(cTabQry)->(DbSkip())

		End

		(cTabQry)->(DbCloseArea())

	Next y

	oFwMsEx:Activate()

	cArq := "PS_"+DTOS(Date())+"_"+StrTran(Time(),":","")+".xml"

	MsgRun( "Gerando o arquivo, aguarde...", "Posição de Estoque customizado", {|| oFwMsEx:GetXMLFile( cArq ) } )

	If !ExistDir( cCaminho )
		MakeDir( cCaminho)
	Endif


	If __CopyFile( cArq, cCaminho + cArq )

		If !ApOleClient("MSExcel")
			MsgAlert("Microsoft Excel não instalado!")
		Else
			oExcelApp := MsExcel():New()
			oExcelApp:WorkBooks:Open( cCaminho + cArq )
			oExcelApp:SetVisible(.T.)
		EndIF

	Else
		MsgInfo( "Arquivo não copiado para temporário do usuário." )
	Endif

Return

Static Function ConvCpo(_cCampo,_cConteudo)

	Local cRet := ""
	Local nX
	Local cDe  := ""
	Local cAte := ""

	If _cCampo == "B2_COD"

		If "-" $ _cConteudo

			cDe  := Substr(_cConteudo,1,At(_cConteudo,"-")-1)
			cAte := Substr(_cConteudo,At(_cConteudo,"-")+1,TamSx3("B2_COD")[1])
			cRet := " AND B2_COD BETWEEN '"+cDe+"' AND '"+cAte+"' "

		ElseiF ";" $ _cConteudo
			
            cRet += " AND B2_COD IN ('"

			For nX := 1 To Len(_cConteudo)

				If Substr(_cConteudo,nX,1) == ";"
					cRet += "','"
				Endif
				
                cRet += Substr(_cConteudo,nX,1)

			Next nX

			cRet += "') "
		Else
			cRet := " AND B2_COD = '"+_cConteudo+"' "
		Endif

	ElseIf _cCampo == "B2_LOCAL"
		
        If "-" $ _cConteudo

			cDe  := Substr(_cConteudo,1,At(_cConteudo,"-")-1)
			cAte := Substr(_cConteudo,At(_cConteudo,"-")+1,TamSx3("B2_LOCAL")[1])
			cRet := " AND B2_LOCAL BETWEEN '"+cDe+"' AND '"+cAte+"' "

		ElseiF ";" $ _cConteudo
			
            cRet += " AND B2_LOCAL IN ('"

			For nX := 1 To Len(_cConteudo)

				If Substr(_cConteudo,nX,1) == ";"
					cRet += "','"
				Endif
				
                cRet += Substr(_cConteudo,nX,1)

			Next nX

			cRet += "') "
		Else
			cRet := " AND B2_LOCAL = '"+_cConteudo+"' "
		Endif
	
    ElseIf _cCampo == "B1_GRUPO"

		If "-" $ _cConteudo

			cDe  := Substr(_cConteudo,1,At(_cConteudo,"-")-1)
			cAte := Substr(_cConteudo,At(_cConteudo,"-")+1,TamSx3("B1_GRUPO")[1])
			cRet := " AND B1_GRUPO BETWEEN '"+cDe+"' AND '"+cAte+"' "

		ElseiF ";" $ _cConteudo
			
            cRet += " AND B1_GRUPO IN ('"

			For nX := 1 To Len(_cConteudo)

				If Substr(_cConteudo,nX,1) == ";"
					cRet += "','"
				Endif
				
                cRet += Substr(_cConteudo,nX,1)

			Next nX

			cRet += "') "
		Else
			cRet := " AND B1_GRUPO = '"+_cConteudo+"' "
		Endif

	Endif

Return (cRet)
