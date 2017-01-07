<%

'----------------------------------------------------------------
'	Results:
'
'	1- Ótimo			GREEN
'	2- Bom				NAVY
'	3- Alerta			YELLOW
'	4- Ruim				BROW
'	5- Crítico			RED BLINK	  Gauge
Const RES_NONE			= 0			 '   0
Const RES_OTIMO			= 1			 '  20
Const RES_BOM			= 2			 '  40
Const RES_ALERTA		= 3			 '  60
Const RES_RUIM			= 4			 '  80
Const RES_CRITICO		= 5			 ' 100


'----------------------------------------------------------------
Const TSOL_CODI001		= "001"	' Inclusão de Serviço no COM/EO
Const TSOL_CODI015		= "015"	' Incl. de Srv no COM/EO + Audit. Fiscal.

Const TSOL_CODI004		= "004" ' Aceitação de Manuais (MOM, MCQ e MPI)
Const TSOL_CODI012		= "012"	' Aceitação de Suplemento
Const TSOL_CODI016		= "016"	' Aprovação de Programa de Treinamento
Const TSOL_CODI017		= "017"	' Aceitação do Plano de Implementação SGSO

Const TSOL_CODI005		= "005"	' Certificação Inicial
Const TSOL_CODI002		= "002"	' Alteração Dados COM/EO
Const TSOL_CODI013		= "013"	' Solicita Suspensão/Cancelamento
Const TSOL_CODI006		= "006"	' Aceitação de Lista de Capacidade
Const TSOL_CODI007		= "007"	' Auditoria de Fiscalização'

Const TSOL_CODI003		= "003"	' Solicitação de Parecer Técnico
Const TSOL_CODI008		= "008"	' MNT Fora de Sede
Const TSOL_CODI009		= "009"	' Execução de Serviço Excepcional
Const TSOL_CODI014		= "014"	' Cadastramento RT e GR

Const TSOL_CODI018		= "018"	' RCA

Const TST_CODI			= "027"	' Em Distribuição

' TX
Const TSOL_TXCODI017	= "017"	' Inclusão de Novo Modelo de Aeronave
Const TSOL_TXCODI018	= "018"	' Inclusão de Nova Operação na EO
Const TSOL_TXCODI019	= "019"	' Alteração de Capacidade de Manutenção na Base
Const TSOL_TXCODI020	= "020"	' Exclusão de Aeronave da EO
Const TSOL_TXCODI022	= "022"	' Inclusão de Aeronave na EO
Const TSOL_TXCODI009	= "009"	' Certificação Inicial

Const TSOL_TXCODI001	= "001"	' Aceitação de MGM
Const TSOL_TXCODI002	= "002"	' Aceitação/Aprov. de Outros Manuais
Const TSOL_TXCODI004	= "004"	' Aprovação de MEL
Const TSOL_TXCODI005	= "005"	' Aprovação de PM

Const TSOL_TXCODI007	= "007"	' Autorização Excepcional
Const TSOL_TXCODI008	= "008"	' Cadastramento Responsável Técnico
Const TSOL_TXCODI023	= "023"	' Auditoria PTA
Const TSOL_TXCODI006	= "006"	' Auditoria por Demanda

Const TSOL_TXCODI024	= "024"	' RCA

Const TST_TXCODI		= "014"	' Em Distribuição

 %>
