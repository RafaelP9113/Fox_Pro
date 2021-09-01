********************************************************************************************************************* 
* Maia - 16-04-19 - Criacao a exportação para fornecedores

define class obj_entrada as custom
	*- Nome do metodo/função que os objetos linx vão chamar.

	procedure metodo_usuario
		*- Parametros do metodo:
		*- Xmetodo= nome do metodo
		*- Xobjeto= variavel com a referencia ao objeto
		*- Xnome_obj  = nome do objeto
		lparam xmetodo, xobjeto ,xnome_obj
		
		******************** Metodos chamados pelo FORMSET
		*	USR_INIT  												=> NA INICIALIZACAO DA TELA 
		*	USR_ALTER_BEFORE  -> Return .f. Para o Metodo 			=> ANTES DA ALTERACAO ,AO CLICKAR ANTES DE LIBERAR OS CAMPOS
		*	USR_ALTER_AFTER 										=> APOS LIBERAR OS CAMPOS PARA INCLUSAO   
		*	USR_INCLUDE_AFTER 										=> APOS LIBERAR OS CAMPOS PARA INCLUSAO
		*	USR_SEARCH_BEFORE -> Return .f. Para o Metodo PESQUISA	=> ANTES DE FAZER A PESQUISA NO BANCO
		*	USR_SEARCH_AFTER										=> APOS FAZER A PESQUISA NO BANCO
		*	USR_CLEAN_AFTER 										=> APOS LIMPAR A TELA 
		*	USR_REFRESH                                             => 
		*	USR_SAVE_BEFORE   -> Return .f. Para o Metodo 			=> SALVAR ANTES DE JOGAR INFORMACOES NO BANCO
		*	USR_SAVE_AFTER                                          => APOS SALVAR AS INFORMACOES    NO BANCO
		*	USR_ITEN_DELETE_BEFORE -> Return .f. Para o Metodo 		=> ANTES DE EXCLUIR ITEN NA TABELA FILHA '+'
		*	USR_ITEN_DELETE_AFTER                                   => APOS DELETAR UM ITEM NA TABELA FILHA '-' 
		*	USR_ITEN_INCLUDE_BEFORE -> Return .f. Para o Metodo 	=> ANTES DE INCLUIR ITEM NA TABELA FILHA
		*	USR_ITEN_INCLUDE_AFTER									=> APOS INCLUIR ITEM NA TABELA FILHA 
		*
		***************** Metodos que ocorrem dentro da Transaction do Banco de Dados
		*	USR_TRIGGER_AFTER ->Return .f. Para o Salvamento e da Rollback ANTES DE EXECUTAR UMA TRIGGER
		*	USR_TRIGGER_BEFORE ->Return .f. Para o Salvamento e da Rollback


		******************** Metodo chamado pelos Objetos na Validação
		*   USR_VALID -> Return .f. Não deixa o Usuario sair do objeto.

		******************** Metodo chamado pelos PageFrames
		*   USR_ACTIVE_PAGE -> Return .f. Para o Metodo.

		*- Inicio dos procedimentos do Usuario
		*- Testando qual o metodo que esta chamando o metodo entrada

		
		* messagebox( 'Metodo '+ UPPER(xmetodo) + ' Nome Objeto ' + upper(xnome_obj) )

			do case
				case UPPER(xmetodo) == 'USR_INIT'
					SELECT 0
					CREATE CURSOR cur_imp (tipo c(1), rateio c(50), valor c(50), filial c(50))
					thisformset.lx_form_rateios.addobject('bt_imp', 'bt_imp')

				case UPPER(xmetodo) == 'USR_SAVE_BEFORE'
					IF thisformset.p_tool_status = 'I'
						IF !f_vazio(V_ENTRADAS_00.chave_nfe)
							VLC_Select = "select COUNT(1) as num from entradas (nolock) where chave_nfe = '" + V_ENTRADAS_00.chave_nfe + "'"
							F_Select(VLC_Select, 'cur_num')
							
							IF NVL(cur_num.num,0) > 0
								MESSAGEBOX('Chave da nota ja foi usada em outra nota!', 16, wusuario)
								USE IN cur_num
								RETURN .F.
							ENDIF
							
							USE IN cur_num
						ENDIF
					ENDIF
					
																																		
				otherwise
					return .t.				
	
			endcase

	endproc
ENDDEFINE


DEFINE CLASS bt_imp AS botao
	top  = 25
	left = 350
	caption = 'Importar Rateio'
	autosize = .T.
	visible = .T.
	
	PROCEDURE refresh
		this.visible = thisformset.p_tool_status $ 'IA'
	ENDPROC
	
	PROCEDURE click
		SELECT cur_imp
		DELETE ALL
		VLC_File = GETFILE('xls')
		IF !F_vazio(VLC_File)
			VLC_Macro = "append from '" + VLC_File + "' xls"
			&VLC_Macro
		
			SELECT cur_imp
			DELETE FOR EMPTY(tipo)
			DELETE FOR VAL(valor) = 0
			
			SELECT distinct tipo FROM cur_imp INTO CURSOR cur_tipo
			
			SELECT cur_tipo
			SCAN
				IF !(UPPER(ALLTRIM(cur_tipo.tipo)) == 'F' OR UPPER(ALLTRIM(cur_tipo.tipo)) == 'C')
					MESSAGEBOX('A planilha so aceita os tipos F e C!', 16, wusuario)
					USE IN cur_tipo
					RETURN .F.
				ENDIF
			ENDSCAN
			USE IN cur_tipo
			
			SELECT distinct rateio FROM cur_imp WHERE tipo = 'F' INTO CURSOR cur_rf
			select cur_rf
			SCAN
				IF f_vazio(thisformset.lx_form_rateios.tv_cod_matriz_contabil.value)
					MESSAGEBOX('Defina uma matriz contabil!', 16, wusuario)
					USE IN cur_rf
					RETURN .F.
				ENDIF
			
				VLC_select = "select filial from W_FILIAIS where cod_filial = '" + ALLTRIM(cur_rf.rateio) + "' and MATRIZ = 'MATRIZ'"
				F_Select(VLC_Select, 'cur_rateio')
				
				IF EOF('cur_rateio')
					MESSAGEBOX('Rateio ' + ALLTRIM(cur_rf.rateio) + ' nao existe ou esta inativa!', 16, wusuario)
					USE IN cur_rf
					USE IN cur_rateio
					RETURN .F.
				ELSE
					SELECT cur_imp
					replace filial WITH cur_rateio.filial FOR ALLTRIM(rateio) = ALLTRIM(cur_rf.rateio) AND UPPER(ALLTRIM(tipo)) == 'F'
				ENDIF
				
				USE IN cur_rateio
			ENDSCAN
			USE IN cur_rf

			SELECT distinct rateio FROM cur_imp WHERE tipo = 'C' INTO CURSOR cur_cc
			select cur_cc
			SCAN
				VLC_select = "select RATEIO_CENTRO_CUSTO from CTB_CENTRO_CUSTO_RATEIO where RATEIO_CENTRO_CUSTO = '" + ALLTRIM(cur_cc.rateio) + "' and inativo = 0"
				F_Select(VLC_Select, 'cur_rateio')
				
				IF EOF('cur_rateio')
					MESSAGEBOX('Rateio de filial ' + ALLTRIM(cur_cc.rateio) + ' nao existe ou esta inativa!', 16, wusuario)
					USE IN cur_cc
					USE IN cur_rateio
					RETURN .F.
				ENDIF
				USE IN cur_rateio
			ENDSCAN
			USE IN cur_cc		
		ENDIF
		
		VLL_CC = .f.
		VLL_FF = .f.
		VLN_Zerados = 0
		
		SELECT cur_imp
		COUNT FOR VAL(valor) = 0 TO VLN_Zerados
		IF VLN_Zerados > 0
			MESSAGEBOX('Existem valores zerados na planilha de rateios!', 16, wusuario)
			RETURN .f.
		ENDIF
		
*					SET STEP ON 
			
		SCAN
			IF UPPER(ALLTRIM(cur_imp.tipo)) == 'C'
				IF !VLL_CC
					SELECT V_CTB_LANCAMENTO_01_ITEM_RATEIO_CENTRO_CUSTO
					DELETE ALL
					VLL_CC = .T.
				ENDIF

				thisformset.lx_form_rateios.grdRateioCentroCusto.SetFocus()
				thisformset.l_filhas_inclui()

				thisformset.lx_form_rateios.grdRateioCentroCusto.col_tv_CENTRO_CUSTO.tv_CENTRO_CUSTO.p_editado_valor = .t.
				thisformset.lx_form_rateios.grdRateioCentroCusto.col_tv_CENTRO_CUSTO.tv_CENTRO_CUSTO.Value = cur_imp.rateio
				thisformset.lx_form_rateios.grdRateioCentroCusto.col_tv_CENTRO_CUSTO.tv_CENTRO_CUSTO.l_desenhista_recalculo()

				IF thisformset.lx009100_rateio_ccusto12.visible = .T.
					thisformset.lx009100_rateio_ccusto12.cmd_ok.click()
				endif

				thisformset.lx_form_rateios.grdRateioCentroCusto.col_tx_VALOR.tx_VALOR.Value = VAL(cur_imp.valor)
				thisformset.lx_form_rateios.grdRateioCentroCusto.col_tx_VALOR.tx_VALOR.l_desenhista_recalculo()

			ENDIF
			
			IF UPPER(ALLTRIM(cur_imp.tipo)) == 'F'
				IF !VLL_FF
					SELECT V_CTB_LANCAMENTO_01_ITEM_RATEIO_FILIAL
					DELETE ALL
					VLL_FF = .T.
				endif				
			
			*	VLC_Select = "select filial from filiais where cod_filial = '" + cur_imp.rateio + "'"
			*	F_Select(VLC_Select, 'cur_filial')
				thisformset.lx_form_rateios.tv_cod_matriz_contabil.rightclick()

				thisformset.lx_form_rateios.grdRateioFilial.SetFocus()
				thisformset.l_filhas_inclui()

				SELECT V_CTB_LANCAMENTO_01_ITEM_RATEIO_FILIAL
				replace MATRIZ_CONTABIL WITH 'MATRIZ'

				thisformset.lx_form_rateios.grdRateioFilial.col_tv_FILIAL.tV_filial.p_editado_valor = .t.
				thisformset.lx_form_rateios.grdRateioFilial.col_tv_FILIAL.tV_filial.Value = cur_imp.filial
				thisformset.lx_form_rateios.grdRateioFilial.col_tv_FILIAL.tV_filial.valid(.t.)
				thisformset.lx_form_rateios.grdRateioFilial.col_tv_FILIAL.tV_filial.l_desenhista_recalculo()

				IF thisformset.lx009100_rateio_filial12.visible = .T.
					thisformset.lx009100_rateio_filial12.cmd_ok.click()
				endif

				thisformset.lx_form_rateios.grdRateioFilial.col_tx_VALOR.tx_VALOR.Value = VAL(cur_imp.valor)
				thisformset.lx_form_rateios.grdRateioFilial.col_tx_VALOR.tx_VALOR.l_desenhista_recalculo()
			ENDIF


		ENDSCAN

		

		
				
		
	ENDPROC
ENDDEFINE

