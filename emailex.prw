#INCLUDE "TOTVS.CH"
#INCLUDE "PARMTYPE.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "AP5MAIL.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "RESTFUL.CH"
#INCLUDE "APWEBEX.CH"
#INCLUDE "APWEBSRV.CH"
#INCLUDE "SHELL.CH"

/*/{Protheus.doc} U_ENV144()
	Função para enviar e-mail de notifição de notas entregues
	Faturamento -> relatórios -> Notas ficais ->  Notificação NF Não Entregue
	@author Vamilly - Jonivani
	@since 19/01/2023
	@param  MV_ENV144T
			MV_ENV144Vm
/*/

User Function ENV144()
	Local cPergunta   := "Deseja enviar o e-mail de notas não entregue?"
	Local cTitulo     := "ENV144"
	Local nX          := 0
	Local nV          := 0
	Local nEnvTrans   := 1
	Local nEnvVend    := 1
	// Local aPergs      := {}
	Private aDados    := {}
	Private aEmail    := {}
	Private aQtStatus := {}
	Private aQtRegiao := {}
	Private aSeller   := {}
	Private aShipping := {}
	Private cPara     := ""
	Private cAssunto  := ""
	Private cMensagem := ""
	Private cAnexo    := ""
	Private cFrom     := ""
	Private cCC       := ""
	Private cBCC      := ""
	Private cNota     := ""
	Private lJob      :=( Type( "oMainWnd" ) <> "O" ) //(IsBlind())
	Private cENV144t  := Space(60) // MV_ENV144T -> Parametro de e-mail copia para trasnportadora
	Private cENV144v  := Space(60) // MV_ENV144V -> Parametro de e-mail copia por vendedor 
	Private aENV144t  := {}
	Private aENV144v  := {}

	
	If !lJob
		// cPara := FwInputBox("Qual o E-mail que deseja enviar o relatório?", cPara)
		// If FwAlertNoYes("Deseja alterar os parâmetros de E-mail copia?",cTitulo)

		// 	aadd(aPergs, {1, "MV_ENV144T ", cENV144t, "@!", ".T.", "", ".T.", 120, .T.})
		// 	aadd(aPergs, {1, "MV_ENV144V ", cENV144v, "@!", ".T.", "", ".T.", 120, .T.})
			
		// 	If ParamBox(aPergs, "Informe os parâmetros")
		// 		cENV144t := MV_PAR01
		// 		cENV144v := MV_PAR02
		// 		PUTMV("MV_ENV144T", cENV144t)
		// 		PUTMV("MV_ENV144V", cENV144v)
		// 	EndIf
		// EndIF
		 	cENV144t:= GETMV("MV_ENV144T")
			cENV144v:= GETMV("MV_ENV144V")

			aENV144t := Separa(cENV144t,",",.F.)
			aENV144v := Separa(cENV144v,",",.F.)
		If FwAlertNoYes(cPergunta, cTitulo)
			conout("[START-ENV144]-> " + cValToChar(GetTimeStamp(Date())))
			Processa( { || DadosEmail() }, 'Processando dados...', 'Aguarde...')
			// enviar e-mail para cada Transportadora
			For nX:= 1 to Len(aShipping)
				aEmail    := {}
				aQtStatus := {}
				aQtRegiao := {}		
				Processa({||EmailTrans(aShipping[nx][1], AllTrim(aShipping[nx][3]), AllTrim(aShipping[nx][4]))},'Enviando E-mail [' +cValToChar(nEnvTrans)+ '] de [' +cValToChar(Len(aShipping))+ ']','Transportadora: ' + cValToChar(AllTrim(aShipping[nx][1])))
				nEnvTrans++
			Next
			// enviar e-mail para cada vendedor
			For nV:= 1 to Len(aSeller)
				aEmail    := {}
				aQtStatus := {}
				aQtRegiao := {}	
				Processa({||EmailVend(aSeller[nV][3], aSeller[nV][4], aSeller[nV][1])},'Enviando E-mail [' +cValToChar(nEnvVend)+ '] de [' +cValToChar(Len(aSeller))+ ']','Vendedor: ' + cValToChar(AllTrim(aSeller[nV][3]))+ CRLF + "E-mail: " + cValToChar(AllTrim(aSeller[nV][4])))
				nEnvVend++
			Next 
			conout("[END-ENV144]-> " + cValToChar(GetTimeStamp(Date())))	
		EndIF
	Else
		Prepare Environment Empresa '04' Filial '01'
		conout("[START-JOB-ENV144]-> " + cValToChar(GetTimeStamp(Date())))
		DadosEmail()
		cENV144t:= GETMV("MV_ENV144T")
		cENV144v:= GETMV("MV_ENV144V")

		aENV144t := Separa(cENV144t,",",.F.)
		aENV144v := Separa(cENV144v,",",.F.)
		// enviar e-mail para cada transportadora
		For nX:= 1 to Len(aShipping)
			aEmail    := {}
			aQtStatus := {}
			aQtRegiao := {}	
			EmailTrans(aShipping[nx][1], AllTrim(aShipping[nx][3]), AllTrim(aShipping[nx][4]))
		Next
		// enviar e-mail para cada vendedor
		For nV:= 1 to Len(aSeller)
			aEmail    := {}
			aQtStatus := {}
			aQtRegiao := {}	
			EmailVend(aSeller[nV][3], aSeller[nV][4], AllTrim(aSeller[nV][1]))
		Next
		conout("[END-JOB-ENV144]-> " + cValToChar(GetTimeStamp(Date())))
		Reset Environment		
	EndIf	
Return 

/*/{Protheus.doc} DadosEmail
	Função para retronar os dados de trasnportadoras e vendedores
	@author Vamilly - Jonivani
	@since 19/01/2023
/*/
Static Function DadosEmail()
	// Local cAlias := GetNextAliAs()
	Local cTrans := GetNextAliAs()
	Local cVend  := GetNextAliAs()

    // // Lendo Titulos
    // BeginSql AliAs cAlias
	// 	%NoParser%
    //     SELECT 
	// 		  QTY
	// 		, PEDIDO
	// 		, NF
	// 		, EMISSAO_NF
	// 		, DIAS_PARA_FATURAMENTO
	// 		, COD_TRANSP
	// 		, TRANSP
	// 		, PRAZO_MINIMO
	// 		, PRAZO_MAXIMO
	// 		, DATA_ENVIO
	// 		, DATA_ENTREGA
	// 		, DIAS_DA_EMISSAO
	// 		, TEMPO_PARA_ENVIO
	// 		, DIAS_APOS_ENVIO
	// 		, TEMPO_DE_ENTREGA
	// 		, DIAS_DA_ENTREGA
	// 		, STATUS
	// 		, COD_CLIENTE
	// 		, LOJA
	// 		, CLIENTE
	// 		, REGIAO
	// 		, ESTADO
	// 		, REGIAO
	// 		, COD_VENDEDOR
	// 		, VENDEDOR
	// 		, CAPITAL
	// 		, EMISSAO_PEDIDO
	// 	FROM 
	// 		V_REL_NOTAS_ENTREGAS 
	// 	ORDER BY 
	// 		  TRANSP
	// 		, VENDEDOR
	// 		, EMISSAO_NF
    // EndSql

    // //cLQry := GetLastQuery()[2]
    // //MemoWrite("C:\test\selena\sel103geral.sql",cLQry)

    // If Len(cAlias) > 0
    //     (cAlias)->(DbGoTop())
    //     While !(cAlias)->(Eof()) 
	// 		// Salvar todos
	// 		AAdd(aDados,{ (cAlias)->QTY                       ,; //[1]  QTY
	// 					  (cAlias)->PEDIDO                    ,; //[2]  PEDIDO
	// 					  (cAlias)->NF                        ,; //[3]  NF
	// 					  AllTrim((cAlias)->EMISSAO_NF)       ,; //[4]  EMISSAO_NF
	// 					  (cAlias)->DIAS_PARA_FATURAMENTO     ,; //[5]  DIAS_PARA_FATURAMENTO
	// 					  (cAlias)->COD_TRANSP                ,; //[6]  COD_TRANSP
	// 					  (cAlias)->TRANSP                    ,; //[7]  TRANSP
	// 					  (cAlias)->PRAZO_MINIMO              ,; //[8]  PRAZO_MINIMO
	// 					  (cAlias)->PRAZO_MAXIMO              ,; //[9]  PRAZO_MAXIMO
	// 					  AllTrim((cAlias)->DATA_ENVIO)       ,; //[10] DATA_ENVIO
	// 					  AllTrim((cAlias)->DATA_ENTREGA)     ,; //[11] DATA_ENTREGA
	// 					  (cAlias)->TEMPO_PARA_ENVIO          ,; //[12] DIAS_DA_EMISSAO
	// 					  (cAlias)->TEMPO_PARA_ENVIO          ,; //[13] TEMPO_PARA_ENVIO
	// 					  (cAlias)->DIAS_APOS_ENVIO           ,; //[14] DIAS_APOS_ENVIO
	// 					  (cAlias)->TEMPO_DE_ENTREGA          ,; //[15] TEMPO_DE_ENTREGA
	// 					  (cAlias)->DIAS_DA_ENTREGA           ,; //[16] DIAS_DA_ENTREGA
	// 					  (cAlias)->STATUS                    ,; //[17] STATUS
	// 					  (cAlias)->COD_CLIENTE               ,; //[18] COD_CLIENTE
	// 					  (cAlias)->LOJA                      ,; //[19] LOJA
	// 					  (cAlias)->CLIENTE                   ,; //[20] CLIENTE
	// 					  (cAlias)->REGIAO                 ,; //[21] REGIAO
	// 					  (cAlias)->ESTADO                    ,; //[22] ESTADO
	// 					  (cAlias)->REGIAO                    ,; //[23] REGIAO
	// 					  (cAlias)->COD_VENDEDOR              ,; //[24] COD_VENDEDOR
	// 					  (cAlias)->VENDEDOR                  ,; //[25] VENDEDOR
	// 					  (cAlias)->CAPITAL                   ,; //[26] CAPITAL
	// 					  AllTrim((cAlias)->EMISSAO_PEDIDO)   ,; //[27] DT_PEDIDO
	// 					  .F.})
    //         (cAlias)->(Dbskip())
    //     EndDo  

    //     (cAlias)->(DbCloseArea())
    // EndIf

	// Coltetando dados de tranportadora
	BeginSql AliAs cTrans
		SELECT 
			  COD_TRANSP
			, TRANSP
			, COUNT(NF) AS 'QTD_NF'
			, EMAIL_TRANSP
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE
			DATA_ENTREGA IS NULL
		GROUP BY  
			  COD_TRANSP
			, TRANSP
			, EMAIL_TRANSP
		ORDER BY 
			TRANSP
	EndSql

	If Len(cTrans) > 0
		(cTrans)->(DbGoTop())
        While !(cTrans)->(Eof())
			AAdd(aShipping,{(cTrans)->TRANSP       ,;
							(cTrans)->QTD_NF       ,;
							(cTrans)->COD_TRANSP   ,;
							(cTrans)->EMAIL_TRANSP ,;
							.F.})
		(cTrans)->(Dbskip())
        EndDo 
		(cTrans)->(DbCloseArea()) 
	EndIf

	// Coltetando dados de vendedor
	BeginSql AliAs cVend
		SELECT 
			COD_VENDEDOR
			, NOME_VENDEDOR
			, VENDEDOR
			, EMAIL_VENDEDOR
			, COUNT(NF) AS 'QTD_NF'
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE
			DATA_ENTREGA IS NULL
		GROUP BY 
			COD_VENDEDOR
			, NOME_VENDEDOR
			, VENDEDOR
			, EMAIL_VENDEDOR
		ORDER BY 
			VENDEDOR
	EndSql

	If Len(cVend) > 0
		(cVend)->(DbGoTop())
        While !(cVend)->(Eof())
			AAdd(aSeller,{  (cVend)->COD_VENDEDOR    ,; //[1]  COD_VENDEDOR
							(cVend)->NOME_VENDEDOR   ,; //[2]  NOME_VENDEDORD
							(cVend)->VENDEDOR        ,; //[3]  VENDEDOR
							(cVend)->EMAIL_VENDEDOR  ,; //[4]  EMAIL_VENDEDOR
							(cVend)->QTD_NF          ,; //[5]  QTD_NF
						.F.})
		(cVend)->(Dbskip())
        EndDo 
		(cVend)->(DbCloseArea()) 
	EndIf

Return 

/*/{Protheus.doc} EmailTrans
	Função para enviar e-mail para as transportadoras
	@author Vamilly - Jonivani
	@since 19/01/2023
/*/
Static Function EmailTrans(cTransp, cCodTransp, cEnviar)
	Local cAlias   := GetNextAliAs()
	Local cAliasSt := GetNextAliAs()
	Local cAliasCd := GetNextAliAs()	
	Local dDtAux   := ""
    Local dDtPrv   := ""
    Local cFilter  := ""
	Local aQtStatus:= {}
	Local aEntregue:= {}

    // Lendo Notas Ficais
    BeginSql AliAs cAlias
		%NoParser%
        SELECT 
			  PEDIDO
			, EMISSAO_PEDIDO
			, NF
			, DIAS_PARA_FATURAMENTO
			, TRANSP
			, DIAS_DA_ENTREGA
			, CLIENTE
			, MUNICIPIO
			, DATA_ENVIO
			, EMISSAO_NF
			, DIAS_APOS_ENVIO
			, STATUS
			, PRAZO_MAXIMO
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_TRANSP = %Exp:cCodTransp%
			AND DATA_ENTREGA IS NULL
		ORDER BY 
			EMISSAO_NF
    EndSql

    If Len(cAlias) > 0
        (cAlias)->(DbGoTop())
        While !(cAlias)->(Eof()) 
			
			dDtAux   := CTOD((cAlias)->DATA_ENVIO)
 			If !Empty(dDtAux)
				dDtAux := DaySum(dDtAux, (cAlias)->PRAZO_MAXIMO)
				dDtPrv := DTOC(dDtAux)
			Else
				dDtAux   := CTOD((cAlias)->EMISSAO_NF)
				dDtAux := DaySum(dDtAux, (cAlias)->PRAZO_MAXIMO)
				dDtPrv := DTOC(dDtAux)
			EndIf
			
			AAdd(aEmail,{ (cAlias)->PEDIDO                    ,; //[1]  PEDIDO
						  AllTrim((cAlias)->EMISSAO_PEDIDO)   ,; //[2] EMISSAO_PEDIDO
						  (cAlias)->NF                        ,; //[3]  NF
						  AllTrim((cAlias)->EMISSAO_NF)       ,; //[4]  EMISSAO_NF
						  (cAlias)->DIAS_PARA_FATURAMENTO     ,; //[5]  DIAS_PARA_FATURAMENTO
						  (cAlias)->TRANSP                    ,; //[6] TRANSP
						  (cAlias)->DIAS_DA_ENTREGA           ,; //[7] DIAS_DA_ENTREGA
						  (cAlias)->CLIENTE                   ,; //[8] CLIENTE
						  (cAlias)->MUNICIPIO                 ,; //[9] MUNICIPIO
						  AllTrim((cAlias)->DATA_ENVIO)       ,; //[10] DATA_ENVIO
						  AllTrim((cAlias)->EMISSAO_NF)       ,; //[11] EMISSAO_NF
						  (cAlias)->DIAS_APOS_ENVIO           ,; //[12] DIAS_APOS_ENVIO
						  (cAlias)->STATUS                    ,; //[13] STATUS
						  dDtPrv                              ,; //[14] PREVISAO
						  (cAlias)->PRAZO_MAXIMO              ,; //[15] PRAZO_MAXIMO
						  .F.})
            (cAlias)->(Dbskip())
        EndDo  
        (cAlias)->(DbCloseArea())
    EndIf

	BeginSql AliAs cAliasSt
		%NoParser%
		SELECT 
			TRANSP
			, STATUS
			, COUNT(NF) AS 'QTD_NF'
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_TRANSP = %Exp:cCodTransp%
			AND DATA_ENTREGA IS NULL
		GROUP BY  
			COD_TRANSP
			, TRANSP
			, STATUS
		ORDER BY 
			TRANSP
			, STATUS
    EndSql

	If Len(cAliasSt) > 0
        (cAliasSt)->(DbGoTop())
        While !(cAliasSt)->(Eof())			
			AAdd(aQtStatus,{(cAliasSt)->TRANSP   ,; //[1] TRANSP
						  	(cAliasSt)->STATUS   ,; //[2] STATUS
						  	(cAliasSt)->QTD_NF   ,; //[3] QTD_NF
						  	.F.})
            (cAliasSt)->(Dbskip())
        EndDo  
        (cAliasSt)->(DbCloseArea())
    EndIf

	BeginSql AliAs cAliasCd
		%NoParser%
		SELECT 
			  REGIAO
			, COUNT(NF) AS 'QTD_NF'
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_TRANSP = %Exp:cCodTransp%
			AND DATA_ENTREGA IS NULL
		GROUP BY  
			REGIAO
		ORDER BY
			REGIAO
    EndSql

	If Len(cAliasCd) > 0
        (cAliasCd)->(DbGoTop())
        While !(cAliasCd)->(Eof())			
			AAdd(aQtRegiao,{(cAliasCd)->REGIAO,; //[1] REGIAO
						  	(cAliasCd)->QTD_NF   ,; //[2] QTD_NF
						  	.F.})
            (cAliasCd)->(Dbskip())
        EndDo  
        (cAliasCd)->(DbCloseArea())
    EndIf

	If Len(aEmail) > 0

		cNota     := "Relação de Notas Fiscais não entregue, transportadora " + cValToChar(cTransp)
		cAssunto  := "Relatório NF Não Entregue, transportadora " + cValToChar(cTransp)
		cPara     := AllTrim(cEnviar)
		cCC       := cENV144t
		cFilter   := "1"
		cAnexo    := GerAnexo(cFilter, aEmail, aEntregue) // 1- Transportadora 2- vendedor, Dados
		cMensagem := LayoutEmail(aEmail, cNota, aQtStatus, aQtRegiao)
		SendMail144(cPara, cAssunto, cMensagem, cAnexo, cFrom, cCC, cBCC)  
		// Apaga o anexo na pasta (servidor)
		FErase(cAnexo)
	EndIf	
Return

/*/{Protheus.doc} EmailVend
	Função para enviar e-mail para os vendedores
	@author Vamilly - Jonivani
	@since 19/01/2023
/*/
Static Function EmailVend(cVendedor, cEnviar, cCodVend)
	Local cAlias     := GetNextAliAs()
	Local cAliasSt   := GetNextAliAs()
	Local cAliasCd   := GetNextAliAs()
	Local cNfEntr    := GetNextAliAs()
	Local cAliasGR   := GetNextAliAs()
	Local dDtAux     := ""
    Local dDtPrv     := ""
	Local cFilter    := ""
	Local aEntregue  := {}
	Local aGerent    := {}
	Local cGerents   := ""
	Local cNoEnvGer  := SupergetMv("SL_NENV144", .F. , "163") // Gerentes que não querem receber copias. Passar os argumentos por BARRAS EX: 163/200
	Local nA         := 0
	
    // START - Notas ficais não entregue
    BeginSql AliAs cAlias
		%NoParser%
        SELECT 
			  PEDIDO
			, EMISSAO_PEDIDO
			, NF
			, DIAS_PARA_FATURAMENTO
			, TRANSP
			, DIAS_DA_ENTREGA
			, CLIENTE
			, MUNICIPIO
			, DATA_ENVIO
			, EMISSAO_NF
			, DIAS_APOS_ENVIO
			, STATUS
			, PRAZO_MAXIMO
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_VENDEDOR = %Exp:cCodVend%
			AND DATA_ENTREGA IS NULL
		ORDER BY 
			EMISSAO_NF
    EndSql

    If Len(cAlias) > 0
        (cAlias)->(DbGoTop())
        While !(cAlias)->(Eof())  

			dDtAux   := CTOD((cAlias)->DATA_ENVIO)
 			If !Empty(dDtAux)
				dDtAux := DaySum(dDtAux, (cAlias)->PRAZO_MAXIMO)
				dDtPrv := DTOC(dDtAux)
			Else
				dDtAux   := CTOD((cAlias)->EMISSAO_NF)
				dDtAux := DaySum(dDtAux, (cAlias)->PRAZO_MAXIMO)
				dDtPrv := DTOC(dDtAux)
			EndIf
			
			AAdd(aEmail,{ (cAlias)->PEDIDO                    ,; //[1]  PEDIDO
						  AllTrim((cAlias)->EMISSAO_PEDIDO)   ,; //[2] EMISSAO_PEDIDO
						  (cAlias)->NF                        ,; //[3]  NF
						  AllTrim((cAlias)->EMISSAO_NF)       ,; //[4]  EMISSAO_NF
						  (cAlias)->DIAS_PARA_FATURAMENTO     ,; //[5]  DIAS_PARA_FATURAMENTO
						  (cAlias)->TRANSP                    ,; //[6] TRANSP
						  (cAlias)->DIAS_DA_ENTREGA           ,; //[7] DIAS_DA_ENTREGA
						  (cAlias)->CLIENTE                   ,; //[8] CLIENTE
						  (cAlias)->MUNICIPIO                 ,; //[9] MUNICIPIO
						  AllTrim((cAlias)->DATA_ENVIO)       ,; //[10] DATA_ENVIO
						  AllTrim((cAlias)->EMISSAO_NF)       ,; //[11] EMISSAO_NF
						  (cAlias)->DIAS_APOS_ENVIO           ,; //[12] DIAS_APOS_ENVIO
						  (cAlias)->STATUS                    ,; //[13] STATUS
						  dDtPrv                              ,; //[14] PREVISAO
						  (cAlias)->PRAZO_MAXIMO              ,; //[15] PRAZO_MAXIMO
						  .F.})
            (cAlias)->(Dbskip())
        EndDo  
        (cAlias)->(DbCloseArea())
    EndIf
	// END - Notas ficais não entregue

	// Start - Notas Ficais Entregues
	 BeginSql AliAs cNfEntr
		%NoParser%
        SELECT 
			  PEDIDO
			, EMISSAO_PEDIDO
			, NF
			, DIAS_PARA_FATURAMENTO
			, TRANSP
			, DIAS_DA_ENTREGA
			, CLIENTE
			, MUNICIPIO
			, DATA_ENVIO
			, EMISSAO_NF
			, DIAS_APOS_ENVIO
			, STATUS
			, PRAZO_MAXIMO
			, DATA_ENTREGA
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_VENDEDOR = %Exp:cCodVend%
			AND DATA_ENTREGA != ''
		ORDER BY 
			EMISSAO_NF
    EndSql

    If Len(cNfEntr) > 0
        (cNfEntr)->(DbGoTop())
        While !(cNfEntr)->(Eof())
			dDtAux := Nil  
			dDtNov := Nil

			dDtAux   := CTOD((cNfEntr)->DATA_ENVIO)
 			If !Empty(dDtAux)
				dDtAux := DaySum(dDtAux, (cNfEntr)->PRAZO_MAXIMO)
				dDtNov := DTOC(dDtAux)
			Else
				dDtAux := CTOD((cNfEntr)->EMISSAO_NF)
				dDtAux := DaySum(dDtAux, (cNfEntr)->PRAZO_MAXIMO)
				dDtNov := DTOC(dDtAux)
			EndIf
			
			AAdd(aEntregue,{ (cNfEntr)->PEDIDO                    ,; //[1]  PEDIDO
						  	 AllTrim((cNfEntr)->EMISSAO_PEDIDO)   ,; //[2] EMISSAO_PEDIDO
						  	 (cNfEntr)->NF                        ,; //[3]  NF
						  	 AllTrim((cNfEntr)->EMISSAO_NF)       ,; //[4]  EMISSAO_NF
						  	 (cNfEntr)->DIAS_PARA_FATURAMENTO     ,; //[5]  DIAS_PARA_FATURAMENTO
						  	 (cNfEntr)->TRANSP                    ,; //[6] TRANSP
						  	 (cNfEntr)->DIAS_DA_ENTREGA           ,; //[7] DIAS_DA_ENTREGA
						  	 (cNfEntr)->CLIENTE                   ,; //[8] CLIENTE
						  	 (cNfEntr)->MUNICIPIO                 ,; //[9] MUNICIPIO
						  	 AllTrim((cNfEntr)->DATA_ENVIO)       ,; //[10] DATA_ENVIO
						  	 AllTrim((cNfEntr)->EMISSAO_NF)       ,; //[11] EMISSAO_NF
						  	 (cNfEntr)->DIAS_APOS_ENVIO           ,; //[12] DIAS_APOS_ENVIO
						  	 (cNfEntr)->STATUS                    ,; //[13] STATUS
						  	 dDtNov                               ,; //[14] PREVISAO
						  	 (cNfEntr)->PRAZO_MAXIMO              ,; //[15] PRAZO_MAXIMO
						  	 (cNfEntr)->DATA_ENTREGA              ,; //[16] DATA_ENTREGA
						  	 .F.})
            (cNfEntr)->(Dbskip())
        EndDo  
        (cNfEntr)->(DbCloseArea())
    EndIf
	// End - Notas Ficais Entregues

	BeginSql AliAs cAliasSt
		%NoParser%
		SELECT 
			VENDEDOR
			, STATUS
			, COUNT(NF) AS 'QTD_NF'
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_VENDEDOR = %Exp:cCodVend%
			AND DATA_ENTREGA IS NULL
		GROUP BY  
			VENDEDOR
			, STATUS
		ORDER BY 
			VENDEDOR
			, STATUS
    EndSql

	If Len(cAliasSt) > 0
        (cAliasSt)->(DbGoTop())
        While !(cAliasSt)->(Eof())			
			AAdd(aQtStatus,{(cAliasSt)->VENDEDOR ,; //[1] VENDEDOR
						  	(cAliasSt)->STATUS   ,; //[2] STATUS
						  	(cAliasSt)->QTD_NF   ,; //[3] QTD_NF
						  	.F.})
            (cAliasSt)->(Dbskip())
        EndDo  
        (cAliasSt)->(DbCloseArea())
    EndIf

	BeginSql AliAs cAliasCd
		%NoParser%
		SELECT 
			  REGIAO
			, COUNT(NF) AS 'QTD_NF'
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE 
			COD_VENDEDOR = %Exp:cCodVend%
			AND DATA_ENTREGA IS NULL
		GROUP BY  
			REGIAO
		ORDER BY
			REGIAO
    EndSql

	If Len(cAliasCd) > 0
        (cAliasCd)->(DbGoTop())
        While !(cAliasCd)->(Eof())			
			AAdd(aQtRegiao,{(cAliasCd)->REGIAO,; //[1] REGIAO
						  	(cAliasCd)->QTD_NF   ,; //[2] QTD_NF
						  	.F.})
            (cAliasCd)->(Dbskip())
        EndDo  
        (cAliasCd)->(DbCloseArea())
    EndIf

	cNoEnvGer := '%' + FormatIn(cNoEnvGer, '/') + '%' // Tratamento para usar o IN

	// e-mail copia para gerente de cada vendedor
	BeginSql AliAs cAliasGR
		%NoParser%
		SELECT 
			DISTINCT
			COD_GERENTE
			, EMAIL_GERENTE
		FROM 
			V_REL_NOTAS_ENTREGAS
		WHERE
			DATA_ENTREGA IS NULL
			AND COD_VENDEDOR = %Exp:cCodVend%
			AND COD_GERENTE NOT IN %Exp:cNoEnvGer% 
		GROUP BY 
			COD_GERENTE
			, EMAIL_GERENTE
		ORDER BY 
			COD_GERENTE
    EndSql

	If Len(cAliasGR) > 0
        (cAliasGR)->(DbGoTop())
        While !(cAliasGR)->(Eof())			
			AAdd(aGerent,(cAliasGR)->EMAIL_GERENTE)
            (cAliasGR)->(Dbskip())
        EndDo  
        (cAliasGR)->(DbCloseArea())
		
		For nA:= 1 to Len(aGerent)
			cGerents +=	AllTrim(aGerent[nA])+ ","
		Next
		cGerents := substring(cGerents, 1, Len(cGerents) -1)
    EndIf

	If Len(aEmail) > 0
		cAnexo    := ""
		cNota     := "Relação de Notas Fiscais não entregue vendedor " + cValToChar(cVendedor)
		cAssunto  := "Relatório NF Não Entregue, vendedor " + cValToChar(cVendedor)
		cPara     := AllTrim(cEnviar)
		cCC       := cENV144v + "," + cGerents
		cFilter   := "2"
		cAnexo    := GerAnexo(cFilter, aEmail, aEntregue) // 1- Transportadora 2- vendedor, Dados
		cMensagem := LayoutEmail(aEmail, cNota, aQtStatus, aQtRegiao)
		SendMail144(cPara, cAssunto, cMensagem, cAnexo, cFrom, cCC, cBCC)
		// Apaga o anexo na pasta (servidor)
		FErase(cAnexo)
	EndIf	
Return

/*/{Protheus.doc} LayoutEmail
	Função para montar o html do e-mail
	@author Vamilly - Jonivani
	@since 19/01/2023
/*/
Static Function LayoutEmail(aDados, cNota, aQtStatus, aQtRegiao)
	Local cHtmlCompleto := ""
	Local cHtmlTable    := ""
	Local cHtmlSt       := ""
	Local cHtmlCd       := ""
	Local cBodyTable    := ""
	Local cTableSt      := ""
	Local cTableCd      := ""
	// Homologação:= D:\TOTVS12_TESTE\TOTVS12\Microsiga\Ambientes\teste_mrp\workflow\ENV144
	// Prod       := D:\TOTVS12\Microsiga\Ambientes\Producao\workflow\ENV144
	Local cArqIndex    := "\workflow\ENV144\index.html"
	Local cArqItens    := "\workflow\ENV144\bodyTable.html"
	Local cArqSt       := "\workflow\ENV144\bodyTableStatus.html"
	Local cArqCd       := "\workflow\ENV144\bodyTableCidade.html"
	Local nI           := 0
	Local nX           := 0
	Local nC           := 0
	Local cClassStatus := ""
	Local cDiaFat      := ""
	Local cPrazoMax    := ""
	Local cProg        := ""


	//Carrega Index
    oFileIndex := FWFileReader():New(cArqIndex)
    If (oFileIndex:Open())
    	While (oFileIndex:hasLine())
    	  //Conout(oFile:GetLine())
    	  cHtmlCompleto += oFileIndex:GetLine()
    	EndDo
    	oFileIndex:Close()
    Endif

	// Carrega  Tabela Cidade
	oFileCd := FWFileReader():New(cArqCd)
    If (oFileCd:Open())
		While (oFileCd:hasLine())
		//Conout(oFile:GetLine())		
			If !Empty(cHtmlCompleto)
				cHtmlCd += oFileCd:GetLine()
			EndIF
		EndDo
     oFileCd:Close()
    Endif

	//Carrega tabela de HTML de Status
    oTableSt := FWFileReader():New(cArqSt)
    If (oTableSt:Open())
		while (oTableSt:hasLine())
		//Conout(oFile:GetLine())
		cHtmlSt += oTableSt:GetLine()
		end
		oTableSt:Close()
    Endif

	//Carrega matriz de itens da tabela
    oFileTable := FWFileReader():New(cArqItens)
    If (oFileTable:Open())
		while (oFileTable:hasLine())
		//Conout(oFile:GetLine())
		cHtmlTable += oFileTable:GetLine()
		end
		oFileTable:Close()
    Endif

	cHtmlCompleto := StrTran(cHtmlCompleto,"@TITULO@"  ,  cNota)
	cHtmlCompleto := StrTran(cHtmlCompleto,"@TOTALNF@"  , cValToChar(Len(aDados)))

	// Tabela de Staus
	For nX := 1 to len(aQtStatus)
		
		cTableSt += cHtmlSt
		cTableSt := StrTran(cTableSt, "@NFSTATUS@" , aQtStatus[nX][2] ) //[1] NFSTATUS
		cTableSt := StrTran(cTableSt, "@QTD_NF@"   , cValToChar(aQtStatus[nX][3]) ) //[2] QTD_NFs

	Next
	cHtmlCompleto := StrTran(cHtmlCompleto,"@TABLEST@", cTableSt)

	// Tabela de Regiao
	For nC := 1 to len(aQtRegiao)
		
		cTableCd += cHtmlCd
		cTableCd := StrTran(cTableCd, "@NFREGIAO@"        , aQtRegiao[nC][1] ) //[1] NFCIDADE
		cTableCd := StrTran(cTableCd, "@QTD_NF_REGIAO@"   , cValToChar(aQtRegiao[nC][2]) ) //[2] QTD_NF

	Next
	cHtmlCompleto := StrTran(cHtmlCompleto,"@TABLERG@", cTableCd)

	// Tabela de notas 
	For nI := 1 to len(aDados)
		//cDados := StrTran(cHtmlItDia,"@STATUS@"      ,alltrim(Transform(aDados[i][4],"@E 99,999,999.99"))
		/* STATUS :=   	NÃO COLETADO EM ATRASO;
						EM TRANSITO DENTRO DO PRAZO;
						EM TRANSITO COM ATRASO;
						ENTREGUE COM ATRASO;
						FINALIZADO PRAZO;
						NÃO COLETADO
		*/

		If AllTrim(aDados[nI][13]) == "EM SEPARAÇÃO"
			cClassStatus := "#6c757d"	
		ElseIf AllTrim(aDados[nI][13]) == "NÃO COLETADO EM ATRASO"
			cClassStatus := "#dc3545"
		ElseIf AllTrim(aDados[nI][13]) == "EM TRANSITO DENTRO DO PRAZO"
			cClassStatus := "#17a2b8"
		ElseIf AllTrim(aDados[nI][13]) == "EM TRANSITO COM ATRASO"
			cClassStatus := "#ffc107"
		ElseIf AllTrim(aDados[nI][13]) == "FINALIZADO PRAZO"
			cClassStatus := "#28a745"	
		EndIF
		
		// (Prazo maximo / tempo apos envio) * 100
		cProg := cValToChar(Round(((aDados[nI][12]/aDados[nI][15])*100), 0 ))+ "%"
		cDiaFat := cValToChar(aDados[nI][5]) + " Dias"

		If !Empty(aDados[nI][15]) //[15]PRAZO_MAXIMO
			cPrazoMax := cValToChar(aDados[nI][15]) + " Dias"	
		Else
			cPrazoMax := ""	
		EndIf

		cBodyTable += cHtmlTable
		cBodyTable := StrTran(cBodyTable, "@PEDIDO@"      ,            aDados[nI][1] ) //[1] PEDIDO
		cBodyTable := StrTran(cBodyTable, "@DATA_PDP@"    ,            aDados[nI][2] ) //[2] EMISSAO_PEDIDO
		cBodyTable := StrTran(cBodyTable, "@NF@"          ,            aDados[nI][3] ) //[3] NF
		cBodyTable := StrTran(cBodyTable, "@DATA_NF@"     ,            aDados[nI][4] ) //[4] EMISSAO_NF
		cBodyTable := StrTran(cBodyTable, "@TP_FAT@"      ,                  cDiaFat ) //[5] DIAS_PARA_FATURAMENTO
		cBodyTable := StrTran(cBodyTable, "@TRANSP@"      ,            aDados[nI][6] ) //[6] TRANSP
		cBodyTable := StrTran(cBodyTable, "@TP_ENTREGA@"  ,                cPrazoMax ) 
		cBodyTable := StrTran(cBodyTable, "@CLIENTE@"     ,            aDados[nI][8] ) //[8] CLIENTE
		cBodyTable := StrTran(cBodyTable, "@CIDADE@"      ,            aDados[nI][9] ) //[9] MUNICIPIO
		cBodyTable := StrTran(cBodyTable, "@DATA_ENVIO@"  ,            aDados[nI][10]) //[10] DATA_ENVIO
		cBodyTable := StrTran(cBodyTable, "@PREVISAO@"    ,cValToChar(aDados[nI][14])) //[14] PREVISAO
		cBodyTable := StrTran(cBodyTable, "@CLASSSPROG@"  , 		     cClassStatus)
		cBodyTable := StrTran(cBodyTable, "@PROG@"        ,                    cProg )
		cBodyTable := StrTran(cBodyTable, "@PROGRESS@"    ,                	   cProg ) 
		cBodyTable := StrTran(cBodyTable, "@STATUS@"      ,            aDados[nI][13]) //[13] STATUS
		cBodyTable := StrTran(cBodyTable, "@CLASSSTATUS@" , 		     cClassStatus)
		
	Next
						
	cHtmlCompleto := StrTran(cHtmlCompleto,"@BODYTABLE@", cBodyTable)
	
Return cHtmlCompleto

/*/{Protheus.doc} SendMail144
	Função para enviar e-mail
	@author Vamilly - Jonivani
	@since 19/01/2023
/*/
Static Function SendMail144(cPara, cAssunto, cMensagem, cAnexo,cFrom, cCC, cBCC)
	Local oServer     := Nil
	Local oMessage    := Nil
	Local nErr        := 0
	Local cSMTPAddr   := GetMV("MV_WFSMTP")         // Endereço SMTP "smtp.gmail.com" 
	Local cPopAddr    := GetMV("MV_WFPOP3")	        // Pop3 "pop.gmail.com"
	Local cPOPPort    := GetMV("MV_PORPOP3")        // Porta do servidor POP  110  
	Local cSMTPPort   := GetMV("MV_PORSMTP")        // Porta do servidor SMTP 465  
	Local nSMTPTime   := GetMV("MV_RELTIME")        // Timeout SMTP           120 
	Local cUser       := GetMV("MV_RELAUSR")        // Usuario -> reports@selenasulamericana.com.br 
	Local cPass       := "!Selena@2020*" //GetMV("MV_RELPSW")         // Senha   -> !Selena@2020*  
    Local lTest       := GetMV("MV_XTESTNF", , .T.) //[.T.] Envia para e-mail controlado.,[.F.] Envia e-mail para email para usuarios do pedido e solicitação de compras.
	Default cPara     := ""
	Default cCC       := ""
	Default cBCC      := ""
	Default cAssunto  := ""
	Default cMensagem := ""
	Default cAnexo    := "" // O caminho desse arquivo tem que estar dentro de qualquer pasta do ROOTPATH (Normalmente Protheus_Data)

	lTest := .F.

	If Empty(cPara)
		conout("[ERROR]Email destino não definido.")
		If !lJob
			Alert("[ERROR]Email destino não definido.")
		Endif

		return .F.
	EndIf

	// Instancia um novo TMailManager
	oServer := tMailManager():New()

	// Usa SSL na conexao
	// oServer:setUseSSL(.T.)
	conout("[SendEmail]->SSL")
    oServer:SetUseSSL (.T.)
	conout("[SendEmail]->TLS")
    oServer:SetUseTls (.T.)
	conout("[SendEmail]->SSL->TLS")


	// Inicializao
	oServer:init(cPopAddr, cSMTPAddr, cUser, cPass, cPOPPort, cSMTPPort)

	// Alert("Usuario -> "+ cValToChar(cUser)+ ", Senha -> " + cValToChar(cPass))

	// Define o Timeout SMTP
	if oServer:SetSMTPTimeout(nSMTPTime) != 0
		conout("[ERROR] Falha ao definir timeout")
		If !lJob
			Alert("[ERROR] Falha ao definir timeout")
		Endif
		return .F.
	endif

	// Conecta ao servidor
	nErr := oServer:smtpConnect()
	if nErr <> 0
		conOut("[ERROR] Falha ao conectar: " + oServer:getErrorString(nErr))
		If !lJob
        	Alert("[ERROR] Falha ao conectar: " + oServer:getErrorString(nErr))
		EndIf
		oServer:smtpDisconnect()
		return .F.
	endif

	// Realiza autenticacao no servidor
	nErr := oServer:smtpAuth(cUser, cPass)

	if nErr <> 0
		conOut("[ERROR] Falha ao autenticar: " + oServer:getErrorString(nErr))
		If !lJob
        	Alert("[ERROR] Falha ao autenticar: " + oServer:getErrorString(nErr))
		EndIf
		oServer:smtpDisconnect()
		return .F.
	endif

    // Se for teste envia para e-mail controlado
    If lTest
        cPara := "jonivani.pereira@vamilly.com.br"
        cCC   := "jonivani@gmail.com"
        cBCC  := "jonivani_ricardo@hotmail.com"
    EndIf


	// Cria uma nova mensagem (TMailMessage)
	oMessage           := tMailMessage():New()
	oMessage:clear()
	oMessage:cFrom     := cUser 
	oMessage:cTo       := cPara
	oMessage:cCC       := cCC
	oMessage:cBCC      := cBCC
	oMessage:cSubject  := cAssunto
	oMessage:cBody     := cMensagem
    oMessage:MsgBodyType("text/html")

    // Alert(  "cFrom: " + cValToChar(cFrom) + CRLF + ;
    //         "cPara: " + cValToChar(cPara) + CRLF + ; 
    //         "cCC: "   + cValToChar(cCC)   + CRLF + ;
    //         "cBCC: "   + cValToChar(cBCC) + CRLF + ;
    //         "cAssunto: "   + cValToChar(cAssunto))

	If !Empty(cAnexo)
		xRet := oMessage:AttachFile( cAnexo )
		if xRet < 0
			cMsg := "Could not attach file " + cAnexo
			conout( cMsg )
			return
		endif
	EndIf

	// Envia a mensagem
	nErr := oMessage:send(oServer)
	if nErr <> 0
		conout("[ERROR] Falha ao enviar: " + oServer:getErrorString(nErr))
		If !lJob
        	Alert("[ERROR] Falha ao enviar: " + oServer:getErrorString(nErr))
		EndIf
		conout("ENV144->e-mail não enviados: " + cValToChar(cPara))
		oServer:smtpDisconnect()
		return .F.
	Else
		conout("ENV144->e-mail enviados: " + cValToChar(cPara))
	endif

	// Disconecta do Servidor
	oServer:smtpDisconnect()
return .T.

/*/{Protheus.doc} GerAnexo()
	Vai gerar um excel para ser enviado para a transportadora e pra vendedor.
	@type  Function
	@author Jonivani
	@since 18/01/2023
/*/
Static Function GerAnexo(cFilter,aDados, aEntregue)

	// Variáveis de notas ficais não entregues
    Local oTabNEntr  := Nil
	Local aFilNEntr  := {}
    Local cTabNEntr  := ""
    Local cTabNEntr1 := ""
	Local nA         := 0
	Local cNFNENT    := ""
	// Variáveis de notas ficais entregues
    Local oTabEntr   := Nil
	Local aFilEntr   := {}
    Local cTabEntr   := ""
    Local cTabEntr1  := ""
	Local nB         := 0
	Local cNFENT     := ""
	// Cria o diretório temporário dos boletos no server
	// Local cDir       := "D:\TOTVS12\Microsiga\Ambientes\Producao\workflow\ENV144\anexo\"
	Local cDir       := "\workflow\ENV144\anexo\"
	Local cArq       := CriaTrab( , .F. ) + ".xls"
	Local cAnexo     := ""	

	// PEDIDO	DATA PDP	NF	DATA NF	TP FAT	TRANSP	PRZ ETG CONTRADO	CLIENTE	CIDADE	DATA ENVIO	PREVISAO	PROGRESSO	STATUS
	//-----------Tabela Temporaria Para notas não entregues---------//
		/*	aFilNEntr com estrutura de campos:
			[1] Nome
			[2] Tipo
			[3] Tamanho TAMSX3("A1_COD")[1] //Tamanho do campo
			[4] Decimal TAMSX3("A1_COD")[2] //decimal do campo			
		*/

	aadd(aFilNEntr, {"PEDIDO"   , "C", TAMSX3("D2_PEDIDO")[1]   , TamSx3("D2_PEDIDO")[2]})
	aadd(aFilNEntr, {"DATAPDP"  , "C", TAMSX3("C5_EMISSAO")[1]+2, TamSx3("C5_EMISSAO")[2]})
	aadd(aFilNEntr, {"NF"       , "C", TAMSX3("F2_DOC")[1]      , TamSx3("F2_DOC")[2]})
	aadd(aFilNEntr, {"DATANF"   , "C", TAMSX3("F2_EMISSAO")[1]+2, TamSx3("F2_EMISSAO")[2]})
	aadd(aFilNEntr, {"TRANSP"   , "C", TAMSX3("A4_NREDUZ")[1]   , TamSx3("A4_NREDUZ")[2]})
	aadd(aFilNEntr, {"TPFAT"    , "C", TAMSX3("F2_DOC")[1]      , TamSx3("F2_DOC")[2]})
	aadd(aFilNEntr, {"PRZCONTR" , "C", TAMSX3("F2_DOC")[1]      , TamSx3("F2_DOC")[2]})
	aadd(aFilNEntr, {"CLIENTE"  , "C", TAMSX3("A1_NOME")[1]     , TamSx3("A1_NOME")[2]})
	aadd(aFilNEntr, {"CIDADE"   , "C", TAMSX3("A1_MUN")[1]      , TamSx3("A1_MUN")[2]})
	aadd(aFilNEntr, {"DATAENVIO", "C", TAMSX3("F2_DTENVI")[1] +2, TamSx3("F2_DTENVI")[2]})
	aadd(aFilNEntr, {"PREVISAO" , "C", TAMSX3("F2_DTENVI")[1] +2, TamSx3("F2_DTENVI")[2]})
	aadd(aFilNEntr, {"STATUS"   , "C", TAMSX3("A1_NOME")[1]     , TamSx3("A1_NOME")[2]})

	// Start - Cria tabela TMP Notas não entregue
	cTabNEntr  := GetNextAlias()
	oTabNEntr  := FWTemporaryTable():New(cTabNEntr, aFilNEntr)
	oTabNEntr:Create()
	cTabNEntr  := oTabNEntr:GetAlias()
	cTabNEntr1 := '%' + oTabNEntr:GetRealName() + '%'

	//Inserção de dados na tabela TMP 
	DbSelectArea((cTabNEntr))
	(cTabNEntr)->(DbGoTop())           
	For nA := 1 to len(aDados)
		RecLock((cTabNEntr), .T.)
			(cTabNEntr)->PEDIDO   := aDados[nA][1]
			(cTabNEntr)->DATAPDP  := aDados[nA][2]
			(cTabNEntr)->NF       := aDados[nA][3]
			(cTabNEntr)->DATANF   := aDados[nA][4]
			(cTabNEntr)->TPFAT    := cValToChar(aDados[nA][5]) + " Dias"
			(cTabNEntr)->TRANSP   := aDados[nA][6]
			(cTabNEntr)->PRZCONTR := IIf(Empty(aDados[nA][15]),"-", cValToChar(aDados[nA][15]) + " Dias")
			(cTabNEntr)->CLIENTE  := aDados[nA][8]
			(cTabNEntr)->CIDADE   := aDados[nA][9]
			(cTabNEntr)->DATAENVIO:= aDados[nA][10]
			(cTabNEntr)->PREVISAO := aDados[nA][14]
			(cTabNEntr)->STATUS   := aDados[nA][13]
		(cTabNEntr)->(MSUnLock()) 
	Next nA        
	
	cNFNENT := GetNextAlias()  
	
	BeginSQL Alias cNFNENT
		%NoParser%
		SELECT 
				PEDIDO           
			, DATAPDP         
			, NF               
			, DATANF          
			, TPFAT           
			, PRZCONTR
			, CLIENTE          
			, CIDADE           
			, DATAENVIO       
			, PREVISAO         
			, STATUS           
		FROM 
			%Exp:cTabNEntr1% AS NFNENT
		ORDER BY
				DATANF
			, STATUS  
	EndSQL

	(cNFNENT)->(DbGoTop())
	// End - Cria tabela TMP Notas não entregue

	// Adicona array para gerar planilha
	aResult := {}
	aadd(aResult, {cNFNENT, "Nao Entregue"})

	If cFilter = "1" // Transportadora
		cAnexo := ExcelEmail(aResult,.F.,cDir,cArq)
		(cNFNENT)->(DbCloseArea()) 
	ElseIf cFilter = "2" // Vendedor

		aadd(aFilEntr, {"PEDIDO"   , "C", TAMSX3("D2_PEDIDO")[1]   , TamSx3("D2_PEDIDO")[2]})
		aadd(aFilEntr, {"DATAPDP"  , "C", TAMSX3("C5_EMISSAO")[1]+2, TamSx3("C5_EMISSAO")[2]})
		aadd(aFilEntr, {"NF"       , "C", TAMSX3("F2_DOC")[1]      , TamSx3("F2_DOC")[2]})
		aadd(aFilEntr, {"DATANF"   , "C", TAMSX3("F2_EMISSAO")[1]+2, TamSx3("F2_EMISSAO")[2]})
		aadd(aFilEntr, {"TRANSP"   , "C", TAMSX3("A4_NREDUZ")[1]   , TamSx3("A4_NREDUZ")[2]})
		aadd(aFilEntr, {"TPFAT"    , "C", TAMSX3("F2_DOC")[1]      , TamSx3("F2_DOC")[2]})
		aadd(aFilEntr, {"PRZCONTR" , "C", TAMSX3("F2_DOC")[1]      , TamSx3("F2_DOC")[2]})
		aadd(aFilEntr, {"CLIENTE"  , "C", TAMSX3("A1_NOME")[1]     , TamSx3("A1_NOME")[2]})
		aadd(aFilEntr, {"CIDADE"   , "C", TAMSX3("A1_MUN")[1]      , TamSx3("A1_MUN")[2]})
		aadd(aFilEntr, {"DATAENVIO", "C", TAMSX3("F2_DTENVI")[1] +2, TamSx3("F2_DTENVI")[2]})
		aadd(aFilEntr, {"PREVISAO" , "C", TAMSX3("F2_DTENVI")[1] +2, TamSx3("F2_DTENVI")[2]})
		aadd(aFilEntr, {"STATUS"   , "C", TAMSX3("A1_NOME")[1]     , TamSx3("A1_NOME")[2]})
		aadd(aFilEntr, {"ENTREGA"  , "C", TAMSX3("F2_EMISSAO")[1]+2, TamSx3("F2_EMISSAO")[2]})

		// Start - Cria tabela TMP Notas não entregue
		cTabEntr  := GetNextAlias()
		oTabEntr  := FWTemporaryTable():New(cTabEntr, aFilEntr)
		oTabEntr:Create()
		cTabEntr  := oTabEntr:GetAlias()
		cTabEntr1 := '%' + oTabEntr:GetRealName() + '%'

		//Inserção de dados na tabela TMP 
		DbSelectArea((cTabEntr))
		(cTabEntr)->(DbGoTop())           
		For nB := 1 to len(aEntregue)
			RecLock((cTabEntr), .T.)
				(cTabEntr)->PEDIDO   := aEntregue[nB][1]
				(cTabEntr)->DATAPDP  := aEntregue[nB][2]
				(cTabEntr)->NF       := aEntregue[nB][3]
				(cTabEntr)->DATANF   := aEntregue[nB][4]
				(cTabEntr)->TPFAT    := cValToChar(aEntregue[nB][5]) + " Dias"
				(cTabEntr)->TRANSP   := aEntregue[nB][6]
				(cTabEntr)->PRZCONTR := IIf(Empty(aEntregue[nB][15]),"-", cValToChar(aEntregue[nB][15]) + " Dias")
				(cTabEntr)->CLIENTE  := aEntregue[nB][8]
				(cTabEntr)->CIDADE   := aEntregue[nB][9]
				(cTabEntr)->DATAENVIO:= aEntregue[nB][10]
				(cTabEntr)->PREVISAO := aEntregue[nB][14]
				(cTabEntr)->STATUS   := aEntregue[nB][13]
				(cTabEntr)->ENTREGA  := aEntregue[nB][16]
			(cTabEntr)->(MSUnLock()) 
		Next nB       
		
		cNFENT := GetNextAlias()  
		
		BeginSQL Alias cNFENT
			%NoParser%
			SELECT 
				  PEDIDO           
				, DATAPDP         
				, NF               
				, DATANF          
				, TPFAT           
				, PRZCONTR
				, CLIENTE          
				, CIDADE           
				, DATAENVIO       
				, PREVISAO         
				, STATUS   
				, ENTREGA        
			FROM 
				%Exp:cTabEntr1% AS NFENT
			WHERE 
				ENTREGA != ''
			ORDER BY
				  DATANF
				, STATUS  
		EndSQL

		(cNFENT)->(DbGoTop())
		aadd(aResult, {cNFENT, "Entregues"})
		
		cAnexo := ExcelEmail(aResult,.F.,cDir,cArq)
		(cNFENT)->(DbCloseArea()) 
		(cNFNENT)->(DbCloseArea()) 
	EndIf

	
Return cAnexo

/*/{Protheus.doc} ExcelEmail
	Função para gravar arquivo excel
	@author Vamilly - Jonivani
	@since 19/01/2023
/*/
Static Function  ExcelEmail(aDados, lOpen, cDir, cArq)
	Local nI         As Numeric
	Local nX         As Numeric
	Local cField     As Character
	Local nTipo      As Numeric //Tipo da coluna no Excel; 1-General/2-Number/3-Monetario/4-DateTime
	Local aRow       AS Array
	Local aStruct    As Array
	Local cDirArq    AS Character
	Private oExcel   As Object
	Private cSheet   As Character
	Private cTable   As Character
	Private aAreaAnt As Array

	Default cDir    := GetTempPath()
	Default cArq    := CriaTrab( , .F. ) + '.xls'
	Default lOpen   := .T.	

   
	// Chamada sem dados
	If Len(aDados) <= 0 
        cMsg := "Não foi informardo nenhum alias para gerar planilha."
        cSol := "Informar o alias para geração da planilha."
		conout("[ERROR] Falha ao enviar: " + cMsg)
		If !lJob
        	Help(,, "xExcel",, cMsg, 1, 0, Nil, Nil, Nil, Nil, Nil, {cSol})
		EndIf

	EndIf

    // Popula
    oExcel := FWMSExcel():New()

	// Varre os alias
	For nI := 1 To Len(aDados)
		aStruct := (aDados[nI][1])->(DBStruct())
		cSheet  := aDados[nI][2]
		cTable  := aDados[nI][2]

		// Adiciona guia
		oExcel:AddWorkSheet(cSheet)

		// Adiciona tabela 
		oExcel:AddTable(cSheet, cTable)

		// Adiciona Colunas
		For nX := 1 To Len(aStruct)
			
			// Campo
			cField := aStruct[nX][1]

			// Tipo
			Do Case
				Case aStruct[nX][2] == "N"
					nTipo = 2
				Case aStruct[nX][2] == "D"
					nTipo = 4
				OtherWise
					nTipo = 1
			EndCase

			// Adiciona
			oExcel:AddColumn(cSheet, cTable, cField, 1, nTipo)
		Next nX 

		// Preence as linhas	
		(aDados[nI][1])->(DBGoTop())
		While (aDados[nI][1])->(!Eof())

			// Popula linha	
			aRow := {}
			For nX := 1 To Len(aStruct) 
				Aadd(aRow, (aDados[nI][1])->&(aStruct[nX][1])) 
			Next nZ

			// Adiciona linha
			oExcel:AddRow(cSheet, cTable, aRow)

			DBSelectArea(aDados[nI][1])
			(aDados[nI][1])->(DBSkip())
		EndDo		
	Next nI

    // Monta XML
	oExcel:Activate()
	oExcel:GetXMLFile(cDir + cArq)
	
	//Abre arquivo
	If File(cDir + cArq) .And. lOpen
		ShellExecute("open", cDir + cArq, "", "", SW_SHOWNORMAL)
	EndIf

	cDirArq := cDir + cArq
	
Return cDirArq
