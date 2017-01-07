<%
Class cWsSIGAD


    '****************************************************************************************************
    ' ATRIBUTOS 
    '****************************************************************************************************

        Public NameSpace
        Public Wsdl
        Public numResult


    '****************************************************************************************************
    ' M�TODOS P�BLICOS
    '****************************************************************************************************

        '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Gerar protocolo de servi�o.
        ' Par�metros: cpfCnpj, tipoDocumento, identificacao, interessado, assunto, tipoAcesso, 
        '             usuarioOperacao, tipoSuporte.
        ' Retorno: NumeroProtocolo.
	    Public Function ProtocoloIncluir(tipoDocumento, identificacao, dt_protocolo, classificacao, interessado, assunto, _
                                         tipoAcesso, protocoloResposta,txt_arquivo,txt_referencia,byt_arquivo, usuarioOperacao, tipoSuporte)

            ProtocoloIncluir = null

            dim Valor(12,2)

                Valor(0,0) =   "cod_documento_tipo"
                Valor(0,1) =   "int"
                Valor(0,2) =   tipoDocumento

                Valor(1,0) =   "txt_identificacao"
                Valor(1,1) =   "string"
                Valor(1,2) =   identificacao

                Valor(2,0) =   "dt_protocolo"
                Valor(2,1) =   "DateTime"
                Valor(2,2) =   formataData(dt_protocolo)

                Valor(3,0) =   "cod_classificacao"
                Valor(3,1) =   "int"
                Valor(3,2) =   classificacao

                Valor(4,0) =   "cod_interessado"
                Valor(4,1) =   "int"
                Valor(4,2) =   interessado

                Valor(5,0) =   "txt_assunto"
                Valor(5,1) =   "string"
                Valor(5,2) =   assunto

                Valor(6,0) =   "cod_acesso_tipo"
                Valor(6,1) =   "int"
                Valor(6,2) =   tipoAcesso

                Valor(7,0) =   "cod_protocolo_resposta"
                Valor(7,1) =   "int"
                Valor(7,2) =   protocoloResposta

                Valor(8,0) =   "txt_arquivo"
                Valor(8,1) =   "string"
                Valor(8,2) =   txt_arquivo

                Valor(9,0) =   "txt_referencia"
                Valor(9,1) =   "string"
                Valor(9,2) =   txt_referencia

                Valor(10,0) =   "byt_arquivo"
                Valor(10,1) =   "bin.base64"
                Valor(10,2) =   byt_arquivo

                Valor(11,0) =   "cod_usuario_operacao"
                Valor(11,1) =   "int"
                Valor(11,2) =   usuarioOperacao

                Valor(12,0) =   "cod_suporte_tipo"
                Valor(12,1) =   "int"
                Valor(12,2) =   tipoSuporte
            '-- Fim Valor
				

			
            '--- Invocar m�todo do web service
            set ProtocoloIncluir = invocarMetodo("ProtocoloIncluir", Valor)

        End Function

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Pesquisa o c�digo do documento passando o nome dele como par�metro.
        ' Par�metros: nomeDocumento.
        ' Retorno: C�digo do Documento.
        Public Function DocumentoTipoPesquisar(nomeDocumento)

            DocumentoTipoPesquisar = null

            dim Valor(0,2)
                Valor(0,0) =   "txt_documento_tipo"
                Valor(0,1) =   "string"
                Valor(0,2) =   nomeDocumento
              
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set DocumentoTipoPesquisar = invocarMetodo("DocumentoTipoPesquisar", Valor)

	    End Function

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Pesquisa o c�digo do interessado utilizando o n�mero do CPF como par�metro.
        ' Par�metros: tipoPessoa, interessado, cpfCnpj.
        ' Retorno: Interessado.
        Public Function InteressadoPesquisar(tipoPessoa, interessado, cpfCnpj)

            InteressadoPesquisar = null

            dim Valor(2,2)
                Valor(0,0) =   "cod_pessoa_tipo"
                Valor(0,1) =   "string"
                Valor(0,2) =   tipoPessoa

                Valor(1,0) =   "txt_interessado"
                Valor(1,1) =   "string"
                Valor(1,2) =   interessado

                Valor(2,0) =   "txt_cnpj_cpf"
                Valor(2,1) =   "string"
                Valor(2,2) =   cpfCnpj
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set InteressadoPesquisar = invocarMetodo("InteressadoPesquisar", Valor)

	    End Function

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Incluir dados de interessado.
        ' Par�metros: .
        ' Retorno: Status de inclus�o de usu�rio.
        Public Function InteressadoIncluir(tipoPessoa, interessado, cpfCnpj, orgao, tratamento, cargo, responsavel, endereco, bairro, cidade, estado, pais, cep, telefone, fax, email, site)

            InteressadoIncluir = null

            dim Valor(16,2)
                Valor(0,0) =   "cod_pessoa_tipo"
                Valor(0,1) =   "int"
                Valor(0,2) =   tipoPessoa

                Valor(1,0) =   "txt_interessado"
                Valor(1,1) =   "string"
                Valor(1,2) =   interessado

                Valor(2,0) =   "txt_cnpj_cpf"
                Valor(2,1) =   "string"
                Valor(2,2) =   cpfCnpj

                Valor(3,0) =   "txt_orgao"
                Valor(3,1) =   "string"
                Valor(3,2) =   orgao

                Valor(4,0) =   "txt_forma_tratamento"
                Valor(4,1) =   "string"
                Valor(4,2) =   tratamento

                Valor(5,0) =   "txt_cargo"
                Valor(5,1) =   "string"
                Valor(5,2) =   cargo

                Valor(6,0) =   "txt_responsavel"
                Valor(6,1) =   "string"
                Valor(6,2) =   responsavel

                Valor(7,0) =   "txt_endereco"
                Valor(7,1) =   "string"
                Valor(7,2) =   endereco

                Valor(8,0) =   "txt_bairro"
                Valor(8,1) =   "string"
                Valor(8,2) =   bairro

                Valor(9,0) =   "txt_cidade"
                Valor(9,1) =   "string"
                Valor(9,2) =   cidade

                Valor(10,0) =   "cod_estado"
                Valor(10,1) =   "int"
                Valor(10,2) =   estado

                Valor(11,0) =   "cod_pais"
                Valor(11,1) =   "int"
                Valor(11,2) =   pais

                Valor(12,0) =   "txt_cep"
                Valor(12,1) =   "string"
                Valor(12,2) =   cep

                Valor(13,0) =   "txt_telefone"
                Valor(13,1) =   "string"
                Valor(13,2) =   telefone

                Valor(14,0) =   "txt_fax"
                Valor(14,1) =   "string"
                Valor(14,2) =   fax

                Valor(15,0) =   "txt_email"
                Valor(15,1) =   "string"
                Valor(15,2) =   email

                Valor(16,0) =   "txt_site"
                Valor(16,1) =   "string"
                Valor(16,2) =   site
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set InteressadoIncluir = invocarMetodo("InteressadoIncluir", Valor)

        End Function

        '----------------------------------------------------------------------------------------------------
	    ' Descri��o: .
        ' Par�metros: .
        ' Retorno: .
        Public Function ProtocoloPesquisar(protocolo, numero, copia, interessado, documento, orgao)

            ProtocoloPesquisar = null

            dim Valor(5,2)
                Valor(0,0) =   "cod_protocolo"
                Valor(0,1) =   "string"
                Valor(0,2) =   protocolo

                Valor(1,0) =   "txt_numero"
                Valor(1,1) =   "string"
                Valor(1,2) =   numero

                Valor(2,0) =   "cod_copia"
                Valor(2,1) =   "string"
                Valor(2,2) =   copia

                Valor(3,0) =   "txt_interessado"
                Valor(3,1) =   "string"
                Valor(3,2) =   interessado

                Valor(4,0) =   "cod_documento_tipo"
                Valor(4,1) =   "string"
                Valor(4,2) =   documento

                Valor(5,0) =   "cod_orgao_atual"
                Valor(5,1) =   "string"
                Valor(5,2) =   orgao
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloPesquisar = invocarMetodo("ProtocoloPesquisar", Valor)

        End Function

        '----------------------------------------------------------------------------------------------------
	    ' Descri��o: 
        ' Par�metros: 
        ' Retorno: 
        Public Function ProtocoloAutuar(codProtocolo, codUsuario)

            ProtocoloAutuar = null

            dim Valor(1,2)
                Valor(0,0) =   "cod_protocolo"
                Valor(0,1) =   "int"
                Valor(0,2) =   codProtocolo

                Valor(1,0) =   "cod_usuario_operacao"
                Valor(1,1) =   "int"
                Valor(1,2) =   codUsuario
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloAutuar = invocarMetodo("ProtocoloAutuar", Valor)

        End Function

        '----------------------------------------------------------------------------------------------------
	    ' Descri��o: 
        ' Par�metros: 
        ' Retorno: 
        Public Function ProtocoloJuntar(codProtocoloPrincipal, codProtocoloJuntado, codUsuario, codJuntada)

            ProtocoloJuntar = null

            dim Valor(3,2)
                Valor(0,0) =   "cod_protocolo_principal"
                Valor(0,1) =   "int"
                Valor(0,2) =   codProtocoloPrincipal

                Valor(1,0) =   "cod_protocolo_juntado"
                Valor(1,1) =   "int"
                Valor(1,2) =   codProtocoloJuntado

                Valor(2,0) =   "cod_usuario_operacao"
                Valor(2,1) =   "int"
                Valor(2,2) =   codUsuario

                Valor(3,0) =   "cod_juntada"
                Valor(3,1) =   "int"
                Valor(3,2) =   codJuntada
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloJuntar = invocarMetodo("ProtocoloJuntar", Valor)

        End Function


		'-----------------------------------------------------------------------------------------------
		' Hist�rico
		'	M�todo utilizado para retornar um Hist�rico de Processos/Documentos quanto ao seu Tramite.
		'
		' Parametros:
		'	Nome						Dados		Fun��o
		'	--------------------------------------------------------------------------------------------
		'	cod_protocolo				Inteiro		C�digo do protocolo
		'
		' Retorno:
		'	Nome						Dados		Func�o
		'	--------------------------------------------------------------------------------------------
		'	cod_protocolo				Inteiro		C�digo do protocolo
		'	cod_orgao					Inteiro		C�digo do �rg�o onde a opera��o est� sendo realizada
		'	cod_orgao_origem			Inteiro		C�digo do �rg�o de origem do tr�mite
		'	cod_orgao_destino			Inteiro		C�digo do �rg�o de destino do tr�mite
		'	cod_usuario_movimento		Inteiro		C�digo do usu�rio que efetuou o tr�mite
		'	cod_usuario_recebimento		Inteiro		C�digo do usu�rio que recebeu o tr�mite
		'	dt_movimento				Data/Hora	Data/Hora do tr�mite
		'	dt_recebimento				Data/Hora	Data/Hora do recebimento do tr�mite
		'
        Public Function ProtocoloHistorico(codProtocolo)

            ProtocoloHistorico = null

            dim Valor(0,2)
                Valor(0,0) =   "cod_protocolo"
                Valor(0,1) =   "int"
                Valor(0,2) =   codProtocolo
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloHistorico = invocarMetodo("ProtocoloHistorico", Valor)

        End Function


		'----------------------------------------------------------------------------------------------------
		' Descri��o: Verificar par�metros passados para o metodo de invocar web service.
		' Parametros:
		'	Nome						Dados	Fun��o
		'---------------------------------------------------------------------------------------------
		'	cod_protocolo				Inteiro	Chave do registro a ser distribu�do
		'	cod_orgao					Inteiro	C�digo do �rg�o onde a opera��o est� sendo realizada
		'	cod_usuario_distribuidor	Inteiro	C�digo do usu�rio que efetua a opera��o
		'	cod_usuario_recebedor		Inteiro	C�digo do usu�rio que recebe a opera��o
		'	cod_motivo					Inteiro	C�digo do motivo
		'	txt_providencia				Texto	Texto da provid�ncia informada pelo usu�rio
		'
		' Retorno:
		'	Nome	Tipo de Dados	Func�o
		'	--------------------------------------------------------------------------------------------
		'	cod_saida	Inteiro	Valor que indica se a fun��o foi executada corretamente:
		'				 0 Executado com sucesso
		'				-1 O registro j� est� distribuido
		'				-2 O cod_usuario_distribuidor n�o possui perfil no cod_orgao informado
		'				-3 O cod_usuario_recebedor n�o possui perfil no cod_orgao informado
		'
		Public Function ProtocoloDistribuir(codigoProto, codigoOrgao, usuarioDistribuidor, usuarioRecebedor, codigoMotivo, providencia)

            ProtocoloDistribuir = null

            Dim Valor(5,2)
                Valor(0,0) = "cod_protocolo"
                Valor(0,1) = "int"
                Valor(0,2) = codigoProto

                Valor(1,0) = "cod_orgao"
                Valor(1,1) = "int"    
                Valor(1,2) = codigoOrgao

                Valor(2,0) = "cod_usuario_distribuidor"
                Valor(2,1) = "int"
                Valor(2,2) = usuarioDistribuidor

                Valor(3,0) = "cod_usuario_recebedor"
                Valor(3,1) = "int"
                Valor(3,2) = usuarioRecebedor

                Valor(4,0) = "cod_motivo"
                Valor(4,1) = "int"
                Valor(4,2) = codigoMotivo

                Valor(5,0) = "txt_providencia"
                Valor(5,1) = "string"
                Valor(5,2) = providencia
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloDistribuir = invocarMetodo("ProtocoloDistribuir", Valor)

        End Function


		'----------------------------------------------------------------------------------------------------
		' ProtocoloTramitar
		'
		' Entrada:
		'	Nome					Tipo	Fun��o
		'	---------------------------------------------------------------------------------
		'	cod_protocolo			Inteiro	Chave do registro a ter o coment�rio exclu�do
		'	cod_orgao_origem		Inteiro	C�digo do �rg�o/unidade de origem
		'	cod_orgao_destino		Inteiro	C�digo do �rg�o/unidade de destino
		'	cod_motivo				Inteiro	C�digo do motivo da tramita��o
		'	txt_despacho			Texto	Despacho da tramita��o
		'	cod_usuario_movimento	Inteiro	C�digo do usu�rio que efetua a tramita��o
		'	cod_usuario_recebimento	Inteiro	C�digo do usu�rio que recebe a tramita��o
		'	cod_numero_volume		Inteiro	N�mero de volumes tramitados
		'	cod_numero_pagina		Inteiro	N�mero de p�ginas tramitadas
		'	cod_numero_anexo		Inteiro	N�mero de anexos tramitados
		'	cod_prioridade			Inteiro	C�digo da prioridade da tramita��o.
		'	cod_usuario_cuidado		Inteiro	C�digo do usu�rio que � enviado a tramita��o
		'	dt_prazo_resposta		Texto	Prazo de resposta da tramita��o
		'
		' Sa�da:
		'
		' Retorno:
		'	Nome	Tipo de Dados	Func�o
		'	--------------------------------------------------------------------------------------------
		'	cod_saida	Inteiro	Valor que indica se a fun��o foi executada corretamente:
		'				 0 Executado com sucesso
		'				-1 O registro n�o est� localizado na UORG com c�digo passado em cod_orgao_origem
		'				-2 O usu�rio n�o possui perfil na cod_orgao_origem
		'
		Public Function ProtocoloTramitar(codProto, codOrgaoOrig, codOrgaoDest, codMotivo, despacho, codUserMov, codUserRec, codNumVol, codNumPag, codNroAnex, codPrior, codUserCuid, dtPrazo)

            ProtocoloTramitar = null

            Dim Valor(12,2)
                Valor(0,0) = "cod_protocolo"
                Valor(0,1) = "int"
                Valor(0,2) = codProto

                Valor(1,0) = "cod_orgao_origem"
                Valor(1,1) = "int"    
                Valor(1,2) = codOrgaoOrig

                Valor(2,0) = "cod_orgao_destino"
                Valor(2,1) = "int"
                Valor(2,2) = codOrgaoDest

                Valor(3,0) = "cod_motivo"
                Valor(3,1) = "int"
                Valor(3,2) = codMotivo

                Valor(4,0) = "txt_despacho"
                Valor(4,1) = "string"
                Valor(4,2) = despacho

                Valor(5,0) = "cod_usuario_movimento"
                Valor(5,1) = "int"
                Valor(5,2) = codUserMov

                Valor(6,0) = "cod_usuario_recebimento"
                Valor(6,1) = "int"
                Valor(6,2) = codUserRec

                Valor(7,0) = "cod_numero_volume"
                Valor(7,1) = "int"
                Valor(7,2) = codNumVol

                Valor(8,0) = "cod_numero_pagina"
                Valor(8,1) = "int"
                Valor(8,2) = codNumPag

                Valor(9,0) = "cod_numero_anexo"
                Valor(9,1) = "int"
                Valor(9,2) = codNroAnex

                Valor(10,0) = "cod_prioridade"
                Valor(10,1) = "int"
                Valor(10,2) = codPrior

                Valor(11,0) = "cod_usuario_cuidado"
                Valor(11,1) = "int"
                Valor(11,2) = codUserCuid

                Valor(12,0) = "dt_prazo_resposta"
                Valor(12,1) = "string"
                Valor(12,2) = dtPrazo
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloTramitar = invocarMetodo("ProtocoloTramitar", Valor)

        End Function


	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Verificar par�metros passados para o m�todo de invocar web service.
        ' Par�metros: .
        ' Retorno: .
        Public Function UsuarioPesquisar(codigoUsuario, login, nome, orgao)

            UsuarioPesquisar = null

            Dim Valor(3,2)
                Valor(0,0) = "cod_usuario"
                Valor(0,1) = "int"
                Valor(0,2) = codigoUsuario

                Valor(1,0) = "txt_login"
                Valor(1,1) = "string"    
                Valor(1,2) = UCase(login)

                Valor(2,0) = "txt_nome"
                Valor(2,1) = "string"
                Valor(2,2) = nome

                Valor(3,0) = "cod_orgao"
                Valor(3,1) = "string"
                Valor(3,2) = orgao
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set UsuarioPesquisar = invocarMetodo("UsuarioPesquisar", Valor)

        End Function

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Pesquisa o c�digo do documento passando o nome dele como par�metro.
        ' Par�metros: nomeDocumento.
        ' Retorno: C�digo do Documento.
        Public Function UOrgPesquisar(sigla)

            UOrgPesquisar = null

            dim Valor(1,2)
                Valor(0,0) =   "txt_descricao"
                Valor(0,1) =   "string"
                Valor(0,2) =   ""
              
                Valor(0,0) =   "txt_sigla"
                Valor(0,1) =   "string"
                Valor(0,2) =   sigla

            '-- Fim Valor

            '--- Invocar m�todo do web service
            Set UOrgPesquisar = invocarMetodo("UORGPesquisar", Valor)

	    End Function

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Verificar par�metros passados para o m�todo de invocar web service.
        ' Par�metros: .
        ' Retorno: .
        Public Function PaisPesquisar(pais, sigla)

            PaisPesquisar = null

            Dim Valor(1,2)
                Valor(0,0) =   "txt_pais"
                Valor(0,1) =   "string"
                Valor(0,2) =   pais

                Valor(1,0) =   "txt_sigla"
                Valor(1,1) =   "string"
                Valor(1,2) =   sigla
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set PaisPesquisar = invocarMetodo("PaisPesquisar", Valor)

        End Function

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Verificar par�metros passados para o m�todo de invocar web service.
        ' Par�metros: .
        ' Retorno: .
        Public Function EstadoPesquisar(estado, sigla)

            EstadoPesquisar= null

            Dim Valor(1,2)
                Valor(0,0) = "txt_estado"
                Valor(0,1) = "string"
                Valor(0,2) = estado

                Valor(1,0) = "txt_sigla"
                Valor(1,1) = "string"
                Valor(1,2) = sigla
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set EstadoPesquisar = invocarMetodo("EstadoPesquisar", Valor)

        End Function

        Public Function DocumentoTipoRetornaIdentificacao(documentoTipo,codOrgao)
            
            DocumentoTipoRetornaIdentificacao= null

            Dim Valor(1,2)
                Valor(0,0) = "cod_documento_tipo"
                Valor(0,1) = "int"
                Valor(0,2) = documentoTipo

                Valor(1,0) = "cod_orgao"
                Valor(1,1) = "int"
                Valor(1,2) = codOrgao
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set DocumentoTipoRetornaIdentificacao = invocarMetodo("DocumentoTipoRetornaIdentificacao", Valor)

        End Function


        Public Function ProtocoloArquivoDigitalIncluir(codProcotolo,txtArquivo,bytArquivo,codUsuarioOperacao)

            ProtocoloArquivoDigitalIncluir= null

            Dim Valor(3,2)
                Valor(0,0) = "cod_protocolo"
                Valor(0,1) = "int"
                Valor(0,2) = codProcotolo

                Valor(1,0) = "txt_arquivo"
                Valor(1,1) = "string"
                Valor(1,2) = txtArquivo

                Valor(2,0) = "byt_arquivo"
                Valor(2,1) = "bin.base64"
                Valor(2,2) = bytArquivo

                Valor(3,0) = "cod_usuario_operacao"
                Valor(3,1) = "int"
                Valor(3,2) = codUsuarioOperacao
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloArquivoDigitalIncluir = invocarMetodo("ProtocoloArquivoDigitalIncluir", Valor)

        End Function


        Public Function ProtocoloArquivoDigitalConteudo(TxtArquivo)

            ProtocoloArquivoDigitalConteudo = null

            Dim Valor(0,2)
                Valor(0,0) = "txt_arquivo"
                Valor(0,1) = "string"
                Valor(0,2) = TxtArquivo
            '-- Fim Valor

            '--- Invocar m�todo do web service
            set ProtocoloArquivoDigitalConteudo = invocarMetodo("ProtocoloArquivoDigitalConteudo", Valor)

        End Function


        Public Function ProtocoloArquivoDigitalConsultar(codProcotolo)

            ProtocoloArquivoDigitalConsultar= null

            Dim Valor(0,2)
                Valor(0,0) = "cod_protocolo"
                Valor(0,1) = "int"
                Valor(0,2) = codProcotolo

            '-- Fim Valor

            '--- Invocar m�todo do web service
            Set ProtocoloArquivoDigitalConsultar = invocarMetodo("ProtocoloArquivoDigitalConsultar", Valor)

        End Function



    '****************************************************************************************************
    ' M�TODOS PRIVADOS
    '****************************************************************************************************

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: .
        ' Par�metros: .
        ' Retorno: .
        Private Sub carregaConfigWs()

            me.NameSpace = "http://www.ikhon.com.br/ws/"
            me.Wsdl      = urlWSProton

        End Sub

	    '----------------------------------------------------------------------------------------------------
	    ' Descri��o: .
        ' Par�metros: .
        ' Retorno: .
        Private Function invocarMetodo(metodo, valor)

            invocarMetodo = null

            dim oWebService, oRec

            '--- cWs: Carrega classe para consumo do web service
            call carregaConfigWs()
            set oWebService = new cWebService
                oWebService.WSDL = me.Wsdl

            '--- Invocar m�todo do web service
            set invocarMetodo = oWebService.Invocar(me.NameSpace, metodo, valor)
            me.numResult = oWebService.numResult

        End Function

        '----------------------------------------------------------------------------------------------------
	    ' Descri��o: Formatar data para envio ao wsSIGAD.
        ' Par�metros: Data.
        ' Retorno: Data no formato yyyy-mm-aa.
	    Private Function formataData(data)

		    arrData = Split(data,"/")
		    dia = arrData(0)
		    mes = arrData(1)
		    ano = arrData(2)

		    if Len(arrData(0)) = "1" then
			    dia = "0" & arrData(0)
		    end if

		    if Len(arrData(1)) = "1" then
			    mes = "0" & arrData(1)
		    end if

            formataData = ano & "-" & mes & "-" & dia

	    End Function

End Class
%>