<%

'----------------------------------------------------------------
'	Results:
'
'	1- �timo			GREEN
'	2- Bom				NAVY
'	3- Alerta			YELLOW
'	4- Ruim				BROW
'	5- Cr�tico			RED BLINK	  Gauge
Const RES_NONE			= 0			 '   0
Const RES_OTIMO			= 1			 '  20
Const RES_BOM			= 2			 '  40
Const RES_ALERTA		= 3			 '  60
Const RES_RUIM			= 4			 '  80
Const RES_CRITICO		= 5			 ' 100


'----------------------------------------------------------------
Const TSOL_CODI001		= "001"	' Inclus�o de Servi�o no COM/EO
Const TSOL_CODI015		= "015"	' Incl. de Srv no COM/EO + Audit. Fiscal.

Const TSOL_CODI004		= "004" ' Aceita��o de Manuais (MOM, MCQ e MPI)
Const TSOL_CODI012		= "012"	' Aceita��o de Suplemento
Const TSOL_CODI016		= "016"	' Aprova��o de Programa de Treinamento
Const TSOL_CODI017		= "017"	' Aceita��o do Plano de Implementa��o SGSO

Const TSOL_CODI005		= "005"	' Certifica��o Inicial
Const TSOL_CODI002		= "002"	' Altera��o Dados COM/EO
Const TSOL_CODI013		= "013"	' Solicita Suspens�o/Cancelamento
Const TSOL_CODI006		= "006"	' Aceita��o de Lista de Capacidade
Const TSOL_CODI007		= "007"	' Auditoria de Fiscaliza��o'

Const TSOL_CODI003		= "003"	' Solicita��o de Parecer T�cnico
Const TSOL_CODI008		= "008"	' MNT Fora de Sede
Const TSOL_CODI009		= "009"	' Execu��o de Servi�o Excepcional
Const TSOL_CODI014		= "014"	' Cadastramento RT e GR

Const TSOL_CODI018		= "018"	' RCA

Const TST_CODI			= "027"	' Em Distribui��o

' TX
Const TSOL_TXCODI017	= "017"	' Inclus�o de Novo Modelo de Aeronave
Const TSOL_TXCODI018	= "018"	' Inclus�o de Nova Opera��o na EO
Const TSOL_TXCODI019	= "019"	' Altera��o de Capacidade de Manuten��o na Base
Const TSOL_TXCODI020	= "020"	' Exclus�o de Aeronave da EO
Const TSOL_TXCODI022	= "022"	' Inclus�o de Aeronave na EO
Const TSOL_TXCODI009	= "009"	' Certifica��o Inicial

Const TSOL_TXCODI001	= "001"	' Aceita��o de MGM
Const TSOL_TXCODI002	= "002"	' Aceita��o/Aprov. de Outros Manuais
Const TSOL_TXCODI004	= "004"	' Aprova��o de MEL
Const TSOL_TXCODI005	= "005"	' Aprova��o de PM

Const TSOL_TXCODI007	= "007"	' Autoriza��o Excepcional
Const TSOL_TXCODI008	= "008"	' Cadastramento Respons�vel T�cnico
Const TSOL_TXCODI023	= "023"	' Auditoria PTA
Const TSOL_TXCODI006	= "006"	' Auditoria por Demanda

Const TSOL_TXCODI024	= "024"	' RCA

Const TST_TXCODI		= "014"	' Em Distribui��o

 %>
