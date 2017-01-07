<!-- #include virtual = "/inet2/cfg/ConfigSEI.asp" -->
<!-- #include file = "cWsNew.asp" -->
<%
'Option Explicit

'----------------------------------------------------------------
'
'	Class cSEI
'

Class cSEI

	'Declarations

	Public NameSpace
	Public Wsdl
	Public numResult

	Private oCtrlErr					' Error Object
	Private m_ClassEnable				' If class is active

	Private m_IdProcedimento			' Id interno do processo no SEI, ex.: 1210000000774
	Private m_ProcedimentoFormatado		' Número do processo visível para o usuário, ex: 12.1.000000077-4
	Private m_IdDocumento				' Id interno do documento no SEI, ex.: 1140000000872
	Private m_DocumentoFormatado		' Número do documento visível para o usuário, ex.: 0003934
	Private m_LinkAcesso				' Link para acesso ao documento
	Private m_SerieId					' tipo do documento (ver estrutura Serie)
	Private m_SerieNome					' tipo do documento (ver estrutura Serie)
	Private m_Numero					' Número do documento
	Private m_Data						' Data de geração para documentos internos e para documentos externos é a data informada na tela de cadastro
	Private m_UnElaboradoraId			' Dados da unidade que gerou o documento (ver estrutura Unidade)
	Private m_UnElaboradoraSigla
	Private m_UnElaboradoraNome


	'Class Initialization
	Private Sub Class_Initialize()
		Set oCtrlErr = new cCtrlErr
		m_ClassEnable = True
		If UCase(urlWS) = "NONE" Then
			m_ClassEnable = False
		End If
	End Sub

	'Terminate Class
	Private Sub Class_Terminate()
		Set oCtrlErr = Nothing
	End Sub

	' Test if is enabled
	Public Property Get IsEnabled()
		IsEnabled = m_ClassEnable
	End Property


	'****************************************************************************************************
	' MÉTODOS PÚBLICOS
	'****************************************************************************************************


	'-----------------------------------------------------------------------------------------------------
	'	consultarDocumento
	'
	'	Parâmetros de Entrada
	'	SiglaSistema				Valor informado no cadastro do sistema realizado no SEI
	'	IdentificacaoServico		Valor informado no cadastro do serviço realizado no SEI
	'	IdUnidade					Identificador da unidade no SEI (sugere-se que este id seja armazenado em uma tabela auxiliar do sistema cliente).
	'	ProtocoloDocumento			Número do documento visível para o usuário, ex.: 0003934
	'	SinRetornarAndamentoGeracao	S/N - sinalizador para retorno do andamento de geração
	'	SinRetornarAssinaturas		S/N - sinalizador para retorno das assinaturas do documento
	'	SinRetornarPublicacao		S/N - sinalizador para retorno dos dados de publicação
	'
	'	Parâmetros de Saída estrutura RetornoConsultaDocumento
	'	IdProcedimento				Id interno do processo no SEI, ex.: 1210000000774
	'	ProcedimentoFormatado		Número do processo visível para o usuário, ex: 12.1.000000077-4
	'	IdDocumento					Id interno do documento no SEI, ex.: 1140000000872
	'	DocumentoFormatado			Número do documento visível para o usuário, ex.: 0003934
	'	LinkAcesso					Link para acesso ao documento
	'	Serie						Dados do tipo do documento (ver estrutura Serie)
	'	Numero						Número do documento
	'	Data						Data de geração para documentos internos e para documentos externos é a data informada na tela de cadastro
	'	UnidadeElaboradora			Dados da unidade que gerou o documento (ver estrutura Unidade)
	'	AndamentoGeracao			Informações do andamento de geração (ver estrutura Andamento)
	'	Assinaturas					Conjunto de assinaturas do documento (ver estrutura Assinatura) 
	'								Será um conjunto vazio caso não existam informações.
	'	Publicacao					Informações de publicação do documento (ver estrutura Publicacao). 
	'								Será nulo caso não existam informações.
	'
	Public Function ConsultarDocumento( ProtoDoc )

		Dim tConsultaDocumento
		tConsultaDocumento = Array( "IdProcedimento", _
									"ProcedimentoFormatado", _
									"IdDocumento", _
									"DocumentoFormatado", _
									"LinkAcesso", _
									"Serie/IdSerie", _
									"Serie/Nome", _
									"Numero", _
									"Data" )

		ConsultarDocumento		= -1

		m_IdProcedimento		= ""
		m_ProcedimentoFormatado	= ""
		m_IdDocumento			= ""
		m_DocumentoFormatado	= ""
		m_LinkAcesso			= ""
		m_SerieId				= ""
		m_SerieNome				= ""
		m_Numero				= ""
		m_Data					= ""
		m_UnElaboradoraId		= ""
		m_UnElaboradoraSigla	= ""
		m_UnElaboradoraNome		= ""

		If Not m_ClassEnable Then
			m_IdProcedimento		= "150"
			m_ProcedimentoFormatado	= "00058.000018/2016-11"
			m_IdDocumento			= "151"
			m_DocumentoFormatado	= "0000097"
			m_LinkAcesso			= "http://sei-lab.anac.gov.br/sei/controlador.php?acao=procedimento_trabalhar&amp;id_procedimento=150&amp;id_documento=151"
			m_SerieId				= "12"
			m_SerieNome				= "Memorando"
			m_Numero				= "12"
			m_Data					= "11/04/2016"
			m_UnElaboradoraId		= "110000005"
			m_UnElaboradoraSigla	= "GTGI"
			m_UnElaboradoraNome		= "Gerencia Tecnica de Gestao da Informacao"
			ConsultarDocumento = 1
			Exit Function
		End If

		Dim Valor(6,2)

		' parâmetros iniciais:
		Valor(0,0) =   "SiglaSistema"
		Valor(0,1) =   "String"
		Valor(0,2) =   SiglaSistema

		Valor(1,0) =   "IdentificacaoServico"
		Valor(1,1) =   "String"
		Valor(1,2) =   IdentificacaoServico

		Valor(2,0) =   "IdUnidade"
		Valor(2,1) =   "String"
		Valor(2,2) =   ""			' Qualquer uma	- "110000034" SAR - Superintendência de Aeronavegabilidade

		Valor(3,0) =   "ProtocoloDocumento"
		Valor(3,1) =   "String"
		Valor(3,2) =   Right( "0000000" & ProtoDoc, 7)

		Valor(4,0) =   "SinRetornarAndamentoGeracao"
		Valor(4,1) =   "String"
		Valor(4,2) =   "S"

		Valor(5,0) =   "SinRetornarAssinaturas"
		Valor(5,1) =   "String"
		Valor(5,2) =   "S"

		Valor(6,0) =   "SinRetornarPublicacao"
		Valor(6,1) =   "String"
		Valor(6,2) =   "S"

		'-- Fim Valor

		'--- Invocar método do web service
		Dim rsDiv : Set rsDiv = invocarMetodo("consultarDocumento", Valor, tConsultaDocumento)

		Dim res : res = me.numResult
		If res > 0 Then

			m_IdProcedimento		= rsDiv(tConsultaDocumento(0))
			m_ProcedimentoFormatado	= rsDiv(tConsultaDocumento(1))
			m_IdDocumento			= rsDiv(tConsultaDocumento(2))
			m_DocumentoFormatado	= rsDiv(tConsultaDocumento(3))
			m_LinkAcesso			= rsDiv(tConsultaDocumento(4))
			m_SerieId				= rsDiv(tConsultaDocumento(5))
			m_SerieNome				= rsDiv(tConsultaDocumento(6))
			m_Numero				= rsDiv(tConsultaDocumento(7))
			m_Data					= rsDiv(tConsultaDocumento(8))
			'm_UnElaboradoraId		= rsDiv(tConsultaDocumento())
			'm_UnElaboradoraSigla	= rsDiv(tConsultaDocumento())
			'm_UnElaboradoraNome	= rsDiv(tConsultaDocumento())

			ConsultarDocumento = -1

		End If

	End Function

	Public Property Get IdProcedimento()
		IdProcedimento = m_IdProcedimento
	End Property

	Public Property Get ProcedimentoFormatado()
		ProcedimentoFormatado = m_ProcedimentoFormatado
	End Property

	Public Property Get IdDocumento()
		IdDocumento = m_IdDocumento
	End Property

	Public Property Get DocumentoFormatado()
		DocumentoFormatado = m_DocumentoFormatado
	End Property

	Public Property Get LinkAcesso()
		LinkAcesso = m_LinkAcesso
	End Property

	Public Property Get SerieId()
		SerieId = m_SerieId
	End Property

	Public Property Get SerieNome()
		SerieNome = m_SerieNome
	End Property

	Public Property Get Numero()
		Numero = m_Numero
	End Property

	Public Property Get Data()
		Data = m_Data
	End Property








'''''''''''''''''''''''''''''''''''
'consultarProcedimento
'
'Parâmetros de Entrada
'SiglaSistema	Valor informado no cadastro do sistema realizado no SEI
'IdentificacaoServico	Valor informado no cadastro do serviço realizado no SEI
'IdUnidade	Identificador da unidade no SEI (sugere-se que este id seja armazenado em uma tabela auxiliar do sistema cliente).
'ProtocoloProcedimento	Número do processo visível para o usuário, ex: 12.1.000000077-4
'SinRetornarAssuntos	S/N - sinalizador para retorno dos assuntos do processo
'SinRetornarInteressados	S/N - sinalizador para retorno de interessados do processo
'SinRetornarObservacoes	S/N - sinalizador para retorno das observações das unidades
'SinRetornarAndamentoGeracao	S/N - sinalizador para retorno do andamento de geração
'SinRetornarAndamentoConclusao	S/N - sinalizador para retorno do andamento de conclusão
'SinRetornarUltimoAndamento	S/N - sinalizador para retorno do último andamento
'SinRetornarUnidadesProcedimentoAberto	S/N - sinalizador para retorno das unidades onde o processo se encontra aberto
'SinRetornarProcedimentosRelacionados	S/N - sinalizador para retorno dos processos relacionados
'SinRetornarProcedimentosAnexados	S/N - sinalizador para retorno dos processos anexados
'Parâmetros de Saída
'parametros	Uma ocorrência da estrutura RetornoConsultaProcedimento
'
'

	Public Function ConsultarProcedimento()

		Dim tConsultaProcedimento
		tConsultaProcedimento = Array("IdProcedimento", "ProcedimentoFormatado", "IdProcedimento", _
					"ProcedimentoFormatado", "LinkAcesso", "Serie/IdSerie", "Serie/Nome", "Numero", "Data", _
					"UnidadeElaboradora/IdUnidade", "UnidadeElaboradora/Sigla", "UnidadeElaboradora/Descricao", _
					"AndamentoGeracao", "Assinaturas", "Publicacao" )

		ConsultarProcedimento = -1
		m_Param = ""

		Dim Valor(12,2)

		' parâmetros iniciais:
		Valor(0,0) =   "SiglaSistema"
		Valor(0,1) =   "String"
		Valor(0,2) =   SiglaSistema

		Valor(1,0) =   "IdentificacaoServico"
		Valor(1,1) =   "String"
		Valor(1,2) =   IdentificacaoServico

		Valor(2,0) =   "IdUnidade"
		Valor(2,1) =   "String"
		Valor(2,2) =   ""				' "110000034" ' SAR - Superintendência de Aeronavegabilidade

		Valor(3,0) =   "ProtocoloProcedimento"
		Valor(3,1) =   "String"
		Valor(3,2) =   "00058000018201611"

		Valor(4,0) =   "SinRetornarAssuntos"
		Valor(4,1) =   "String"
		Valor(4,2) =   "N"

		Valor(5,0) =   "SinRetornarInteressados"
		Valor(5,1) =   "String"
		Valor(5,2) =   "N"

		Valor(6,0) =   "SinRetornarObservacoes"
		Valor(6,1) =   "String"
		Valor(6,2) =   "N"

		Valor(7,0) =   "SinRetornarAndamentoGeracao"
		Valor(7,1) =   "String"
		Valor(7,2) =   "N"

		Valor(8,0) =   "SinRetornarAndamentoConclusao"
		Valor(8,1) =   "String"
		Valor(8,2) =   "N"

		Valor(9,0) =   "SinRetornarUltimoAndamento"
		Valor(9,1) =   "String"
		Valor(9,2) =   "N"

		Valor(10,0)=   "SinRetornarUnidadesProcedimentoAberto"
		Valor(10,1)=   "String"
		Valor(10,2)=   "N"

		Valor(11,0)=   "SinRetornarProcedimentosRelacionados"
		Valor(11,1)=   "String"
		Valor(11,2)=   "N"

		Valor(12,0)=   "SinRetornarProcedimentosAnexados"
		Valor(12,1)=   "String"
		Valor(12,2)=   "N"

		'-- Fim Valor

		'--- Invocar método do web service
		Dim rsDiv
		Set rsDiv = invocarMetodo("consultarProcedimento", Valor, tConsultaProcedimento)

		Dim res : res = me.numResult
		If res > 0 Then

m_Param = rsDiv("parametros")

			ConsultarProcedimento = 1

		End If

	End Function




	Public Function ListarUnidades()

		Dim tUnidade : tUnidade = Array("IdUnidade", "Sigla","Descricao")

		ListarUnidades = -1
		m_Param = ""

		dim Valor(3,2)

		' parâmetros iniciais:
		Valor(0,0) =   "SiglaSistema"
		Valor(0,1) =   "String"
		Valor(0,2) =   SiglaSistema

		Valor(1,0) =   "IdentificacaoServico"
		Valor(1,1) =   "String"
		Valor(1,2) =   IdentificacaoServico

		Valor(2,0) =   "IdTipoProcedimento"
		Valor(2,1) =   "String"
		Valor(2,2) =   ""

		Valor(3,0) =   "IdSerie"
		Valor(3,1) =   "String"
		Valor(3,2) =   ""

		'-- Fim Valor

		'--- Invocar método do web service
		Dim rsDiv
		Set rsDiv = invocarMetodo("listarUnidades", Valor, tUnidade)

		Dim res : res = me.numResult
		If res > 0 Then

m_Param = rsDiv("parametros")

			ListarUnidades = 1
		End If

	End Function

	
	Public Function ListarTiposProcedimento()

		Dim tTipoProcedimento : tTipoProcedimento = Array("IdTipoProcedimento", "Nome")

		ListarTiposProcedimento = -1
		m_Param = ""

		Dim Valor(3,2)

		' parâmetros iniciais:
		Valor(0,0) =   "SiglaSistema"
		Valor(0,1) =   "String"
		Valor(0,2) =   SiglaSistema

		Valor(1,0) =   "IdentificacaoServico"
		Valor(1,1) =   "String"
		Valor(1,2) =   IdentificacaoServico

		Valor(2,0) =   "IdUnidade"
		Valor(2,1) =   "String"
		Valor(2,2) =   ""

		Valor(3,0) =   "IdSerie"
		Valor(3,1) =   "String"
		Valor(3,2) =   ""

		'-- Fim Valor

		'--- Invocar método do web service
		Dim rsDiv
		Set rsDiv = invocarMetodo("listarTiposProcedimento", Valor, tTipoProcedimento)

		Dim res : res = me.numResult
		If res > 0 Then

m_Param = rsDiv("parametros")

			ListarTiposProcedimento = 1
		End If


	End Function


	'  lista Unidades
	Public Property Get GetUnidades()
		GetUnidades = m_Param
	End Property


	'****************************************************************************************************
	' MÉTODOS PRIVADOS
	'****************************************************************************************************

	'----------------------------------------------------------------------------------------------------
	' Descrição: .
	' Parâmetros: .
	' Retorno: .
	Private Sub carregaConfigWs()

		me.NameSpace = urlNSpace
		me.Wsdl      = urlWS

	End Sub

	'----------------------------------------------------------------------------------------------------
	' Descrição: .
	' Parâmetros: .
	' Retorno: .
	Private Function invocarMetodo(metodo, valor, tResult)

		invocarMetodo = null

		Dim oWebService, oRec

		'--- cWs: Carrega classe para consumo do web service
		call carregaConfigWs()
		Set oWebService = new cWebService
		oWebService.WSDL = me.Wsdl

		'--- Invocar método do web service
		Set invocarMetodo = oWebService.Invocar(me.NameSpace, metodo, valor, tResult)

		me.numResult = oWebService.numResult

	End Function


End Class

%>

