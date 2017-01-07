<%

'----------------------------------------------------------------
'
'	Class cAIRStatsData
'
'	Date: 27/07/2014
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------

'----------------------------------------------------------------
'
'	Class cAIRStatsData
'
'	Guarda estatísticas relacionada a cada grupo estatístico cadastrado
'

Class cAIRStatsData

	'Declarations
	Private m_oCtrlErr			' Error Object
	Private m_oDbFDH			' DB Object
	Private m_rsDiv				' rsDiv

	' Variables
	Private m_nDate				' Date
Private m_sSolCodi			' Solicitation


	Private m_strGTAR			' GTAR
	Private m_strRBAC			' RBAC

	Private m_nVarANAC			' Processos com a ANAC (em análise com a ANAC)


Private m_nVarGoalANAC			' Tempo meta da ANAC

Private m_nVarGoalDays			' Tempo meta da ANAC

Private m_nVar15dANAC		' Mais de 15 dias com ANAC

Private m_nVar30dANAC		' Mais de 30 dias com ANAC
Private m_nVar60dANAC		' Mais de 60 dias com ANAC
Private m_nVarMaxANAC		' Maior atraso em dias com a ANAC

Private m_nVarTMAANAC			' Tempo médio com ANAC



	Private m_nVar30dClosed		' Fechados nos últimos 30 dias (32 dias)
Private m_nVarTMAClosed			' Tempo médio processo concluído útimos 30 dias (considera o tempo todo ANAC+Empresa)
Private m_nVarItrClosed			' Taxa média de Iterações por processo fechados últimos 30 dias


	Private m_nVar30dDocs		' Documentos emitidos nos últimos 30 dias (32 dias)
Private m_nVarTMADocs			' Tempo médio de resposta da ANAC para Docs emitidos últimos 30 dias

	Private m_nVarClient		' Processos com a empresa
	Private m_nVarDelayClient	' Atrasados com Empresa (> que pendencia+30 ou 90+30)
	Private m_nVarMaxClient		' Maior atraso em dias com Empresa

	Private m_nVarDelivery		' Em Distribuição
	Private m_nVar7dDelivery	' Em Distribuição a mais de 7 dias
	Private m_nVar14dDelivery	' Em Distribuição a mais de 14 dias
	Private m_nVarMaxDelivery	' Maior atraso em dias em Distribuição

	Private m_tStamp			' Timestamp

	'Class Initialization
	Private Sub Class_Initialize()
		Set m_oCtrlErr	= new cCtrlErr
	End Sub
	Public Default Function construct()
		m_nDate				= Date()
		m_strGTAR			= ""
		m_strRBAC			= ""
		m_sSolCodi			= ""
		m_nVarANAC			= 0
		m_nVar30dANAC		= 0
		m_nVar60dANAC		= 0
		m_nVarMaxANAC		= 0
		m_nVar30dClosed		= 0
		m_nVar30dDocs		= 0
		m_nVarClient		= 0
		m_nVarDelayClient	= 0
		m_nVarMaxClient		= 0
		m_nVarDelivery		= 0
		m_nVar7dDelivery	= 0
		m_nVar14dDelivery	= 0
		m_nVarMaxDelivery	= 0
        set construct = me
    End Function

	'Terminate Class
	Private Sub Class_Terminate()
		Set m_oCtrlErr = Nothing
	End Sub

	'  Get Date
	Public Property Get dtDate()
		dtDate = m_nDate
	End Property
	Public Property Let dtDate( value )
		nDate = value
	End Property

	'  Get sSolCodi
	Public Property Get SCodi()
		SCodi = m_sSolCodi
	End Property
	Public Property Let SCodi( value )
		m_sSolCodi = CInt(value)
		m_sSolCodi = String( 3 - Len(m_sSolCodi), "0" ) & m_sSolCodi
	End Property

	'  Get GTAR
	Public Property Get GTAR()
		GTAR = m_strGTAR
	End Property
	Public Property Let GTAR( value )
		m_strGTAR = UCase(value)
	End Property

	'  Get RBAC
	Public Property Get RBAC()
		RBAC = m_strRBAC
	End Property
	Public Property Let RBAC( value )
		m_strRBAC = value
	End Property

	'  Get nVarANAC			' Processos com a ANAC (em análise com a ANAC)
	Public Property Get ANAC()
		ANAC = m_nVarANAC
	End Property
	Public Property Let ANAC( value )
		m_nVarANAC = value
	End Property

	'  Get nVarDelayANAC		' Mais de 30 dias com ANAC
	Public Property Get ANACDelay()
		ANACDelay = m_nVar30dANAC
	End Property
	Public Property Let ANACDelay( value )
		m_nVar30dANAC = value
	End Property

	'  Get nVar30dANAC			' Mais de 30 dias com ANAC
	Public Property Get ANAC30d()
		ANAC30d = m_nVar30dANAC
	End Property
	Public Property Let ANAC30d( value )
		m_nVar30dANAC = value
	End Property

	'  Get nVar60dANAC			' Mais de 60 dias com ANAC
	Public Property Get ANAC60d()
		ANAC60d = m_nVar60dANAC
	End Property
	Public Property Let ANAC60d( value )
		m_nVar60dANAC = value
	End Property

	'  Get nVarMaxANAC			' Acumulado de dias com ANAC
	Public Property Get ANACMaxDays()
		ANACMaxDays = m_nVarMaxANAC
	End Property
	Public Property Let ANACMaxDays( value )
		m_nVarMaxANAC = value
	End Property

	'  Get nVar30dClosed		' Fechados nos últimos 30 dias (32 dias)
	Public Property Get Closed30d()
		Closed30d = m_nVar30dClosed
	End Property
	Public Property Let Closed30d( value )
		m_nVar30dClosed = value
	End Property

	'  Get nVar30dDocs			' Documentos emitidos nos últimos 30 dias (32 dias)
	Public Property Get Docs30d()
		Docs30d = m_nVar30dDocs
	End Property
	Public Property Let Docs30d( value )
		m_nVar30dDocs = value
	End Property

	'  Get nVarClient			' Processos com a Empresa
	Public Property Get Client()
		Client = m_nVarClient
	End Property
	Public Property Let Client( value )
		m_nVarClient = value
	End Property

	'  Get nVarDelayClient		' Atrasados com Empresa
	Public Property Get ClientDelay()
		ClientDelay = m_nVarDelayClient
	End Property
	Public Property Let ClientDelay( value )
		m_nVarDelayClient = value
	End Property

	'  Get VarMaxEmpresa		' Acumulado de dias com Empresa
	Public Property Get ClientMaxDays()
		ClientMaxDays = m_nVarMaxClient
	End Property
	Public Property Let ClientMaxDays( value )
		m_nVarMaxClient = value
	End Property

	'  Get nVarDelivery		' Em Distribuição
	Public Property Get Delivery()
		Delivery = m_nVarDelivery
	End Property
	Public Property Let Delivery( value )
		m_nVarDelivery = value
	End Property

	'  Get VarDelayDelivery	' Em Distribuição a mais de 7 dias
	Public Property Get DeliveryDelay()
		DeliveryDelay = m_nVar7dDelivery
	End Property
	Public Property Let DeliveryDelay( value )
		m_nVar7dDelivery = value
	End Property

	'  Get nVar7dDelivery		' Em Distribuição a mais de 7 dias
	Public Property Get Delivery7d()
		Delivery7d = m_nVar7dDelivery
	End Property
	Public Property Let Delivery7d( value )
		m_nVar7dDelivery = value
	End Property

	'  Get nVar14dDelivery		' Em Distribuição a mais de 14 dias
	Public Property Get Delivery14d()
		Delivery14d = m_nVar14dDelivery
	End Property
	Public Property Let Delivery14d( value )
		m_nVar14dDelivery = value
	End Property

	'  Get nVarMaxDelivery		' Acumulado de dias em Distribuição
	Public Property Get DeliveryMaxDays()
		DeliveryMaxDays = m_nVarMaxDelivery
	End Property
	Public Property Let DeliveryMaxDays( value )
		m_nVarMaxDelivery = value
	End Property

	'  Get nVarSum				' Total de Processos (ANAC + Empresa + Em Distr.)
	Public Property Get Sum()
		Sum = m_nVarANAC + m_nVarClient + m_nVarDelivery
	End Property

	'  Get m_tStamp				' Timestamp
	Public Property Get tStamp()
		tStamp = m_tStamp
	End Property
	

	' Get
	Private Sub getValues()
		m_nDate				= m_rsDiv( "AIRStats_DATE" )
		m_nVarANAC			= m_rsDiv( "AIRStats_ANAC" )
		m_nVar30dANAC		= m_rsDiv( "AIRStats_30D_ANAC" )
		m_nVar60dANAC		= m_rsDiv( "AIRStats_60D_ANAC" )
		m_nVarMaxANAC		= m_rsDiv( "AIRStats_MAX_ANAC" )
		m_nVar30dClosed		= m_rsDiv( "AIRStats_30D_CLOSED" )
		m_nVar30dDocs		= m_rsDiv( "AIRStats_30D_DOCS" )
		m_nVarClient		= m_rsDiv( "AIRStats_CLIENT" )
		m_nVarDelayClient	= m_rsDiv( "AIRStats_DELAYED_CLIENT" )
		m_nVarMaxClient		= m_rsDiv( "AIRStats_MAX_CLIENT" )
		m_nVarDelivery		= m_rsDiv( "AIRStats_DELIVERY" )
		m_nVar7dDelivery	= m_rsDiv( "AIRStats_7D_DELIVERY" )
		m_nVar14dDelivery	= m_rsDiv( "AIRStats_14D_DELIVERY" )
		m_nVarMaxDelivery	= m_rsDiv( "AIRStats_MAX_DELIVERY" )
		m_tStamp			= m_rsDiv( "AIRStats_TIMESTAMP" )
	End Sub


	Public Sub CleanValues()
		m_nVarANAC			= 0
		m_nVar30dANAC		= 0
		m_nVar60dANAC		= 0
		m_nVarMaxANAC		= 0
		m_nVar30dClosed		= 0
		m_nVar30dDocs		= 0
		m_nVarClient		= 0
		m_nVarDelayClient	= 0
		m_nVarMaxClient		= 0
		m_nVarDelivery		= 0
		m_nVar7dDelivery	= 0
		m_nVar14dDelivery	= 0
		m_nVarMaxDelivery	= 0
	End Sub


	' Import
	Public Sub Import( object )
		m_nDate				= object.dtDate
		m_strGTAR			= object.GTAR
		m_strRBAC			= object.RBAC
		m_sSolCodi			= object.SCodi
		m_nVarANAC			= object.ANAC
		m_nVar30dANAC		= object.ANAC30d
		m_nVar60dANAC		= object.ANAC60d
		m_nVarMaxANAC		= object.ANACMaxDays
		m_nVar30dClosed		= object.Closed30d
		m_nVar30dDocs		= object.Docs30d
		m_nVarClient		= object.Client
		m_nVarDelayClient	= object.ClientDelay
		m_nVarMaxClient		= object.ClientMaxDays
		m_nVarDelivery		= object.Delivery
		m_nVar7dDelivery	= object.Delivery7d
		m_nVar14dDelivery	= object.Delivery14d
		m_nVarMaxDelivery	= object.DeliveryMaxDays
		m_tStamp			= object.tStamp
	End Sub


	' Add
	Public Sub Add( object )
		m_nVarANAC			= m_nVarANAC + object.ANAC
		m_nVar30dANAC		= m_nVar30dANAC + object.ANAC30d
		m_nVar60dANAC		= m_nVar60dANAC + object.ANAC60d
		If object.ANACMaxDays > m_nVarMaxANAC Then m_nVarMaxANAC = object.ANACMaxDays
		m_nVar30dClosed		= m_nVar30dClosed + object.Closed30d
		m_nVar30dDocs		= m_nVar30dDocs + object.Docs30d
		m_nVarClient		= m_nVarClient + object.Client
		m_nVarDelayClient	= m_nVarDelayClient + object.ClientDelay
		If object.ClientMaxDays > m_nVarMaxClient Then m_nVarMaxClient = object.ClientMaxDays
		m_nVarDelivery		= m_nVarDelivery + object.Delivery
		m_nVar7dDelivery	= m_nVar7dDelivery + object.Delivery7d
		m_nVar14dDelivery	= m_nVar14dDelivery + object.Delivery14d
		If object.DeliveryMaxDays > m_nVarMaxDelivery Then m_nVarMaxDelivery = object.DeliveryMaxDays
	End Sub


	' Open
	Public Function Open()

		Set m_oDbFDH = (new cDBAccess)( "FDH" )
		If m_oDbFDH.ErrorNumber < 0 then
			m_oCtrlErr.Import( m_oDbFDH.getObjErr() )
			Open = -1
			Exit Function
		End If

	End Function


	'  Get m_oDbFDH
	Public Property Get oDbFDH()
		Set oDbFDH = m_oDbFDH
	End Property


	' Control
	Public Function Exists( dt )

		'--------------------------------------------------------------
		' Verifica se já foi criado algum registro para a data desejada
		Dim querySQL
		querySQL =	"SELECT * FROM AIRStatistics " &_
					" WHERE AIRStats_DATE=#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "# AND " & _
					"  AIRStats_SOLIC='" & m_sSolCodi & "' AND AIRStats_GTAR='" & m_strGTAR & "' AND " & _
					"  AIRStats_RBAC='" & m_strRBAC & "'"
		Set m_rsDiv = m_oDbFDH.getRecSetRd(querySQL)
		If m_rsDiv Is Nothing Then
			m_oCtrlErr.Import( m_oDbFDH.getObjErr() )
			Exists = -1
			Exit Function
		End If

		If m_rsDiv.Eof Then
			m_rsDiv.Close()
			Exists = 0 ' not found
			Exit Function
		End If

		' timestamp
		Dim tStamp : tStamp = rsDiv( "AIRStats_TIMESTAMP" )
		Dim hLast : hLast = DateDiff("h", tStamp, Now())
		If hLast > 3 Then
			Exists = 0 ' not found - to force recalculation
		Else
			Exists = 1 ' already exists
		End If

		m_rsDiv.Close()

	End Function


	Public Function FetchStart()

		' Verify Arguments
		If m_strGTAR = "" Or _
		   m_strRBAC = "" Or _
		   m_sSolCodi = "" Then
			oCtrlErr.Error = "Invalid Arguments."
			FetchStart = -1
			Exit Function
		End If

		'--------------------------------------------------------------
		' Verifica se já foi criado algum registro para a data desejada
		Dim querySQL
		querySQL =	"SELECT * FROM AIRStatistics " &_
					" WHERE AIRStats_GTAR='" & m_strGTAR & "' AND " & _
					"  AIRStats_RBAC='" & m_strRBAC & "' AND" & _
					"  AIRStats_SOLIC='" & m_sSolCodi & "'" & _
					" ORDER BY AIRStats_DATE"
		Set m_rsDiv = m_oDbFDH.getRecSetRd(querySQL)
		If m_rsDiv Is Nothing Then
			m_oCtrlErr.Import( m_oDbFDH.getObjErr() )
			FetchStart = -1
			Exit Function
		End If
		If m_rsDiv.Eof Then
			m_rsDiv.Close()
			FetchStart = 0
			Exit Function
		End If

		' Vars
		getValues()

		FetchStart = 1

	End Function

	Public Function FetchNext()

		' get next
		m_rsDiv.MoveNext

		If m_rsDiv.Eof Then
			m_rsDiv.Close()
			FetchNext = 0
			Exit Function
		End If

		' Vars
		getValues()

		FetchNext = 1

	End Function

	Public Sub FetchClose()
		m_rsDiv.Close()
	End Sub


	' Get values
	Public Function Read( dt )

		' Verify Arguments
		If m_strGTAR = "" Or _
		   m_strRBAC = "" Or _
		   m_sSolCodi = "" Then
			m_oCtrlErr.Error = "Invalid Arguments."
			Close()
			Read = -1
			Exit Function
		End If

		'--------------------------------------------------------------
		' Verifica se já foi criado algum registro para a data desejada
		Dim querySQL
		querySQL =	"SELECT * FROM AIRStatistics " &_
					" WHERE AIRStats_DATE=#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "# AND " & _
					"  AIRStats_SOLIC='" & m_sSolCodi & "' AND " & _
					"  AIRStats_GTAR='" & m_strGTAR & "' AND " & _
					"  AIRStats_RBAC='" & m_strRBAC & "'"
		Set m_rsDiv = m_oDbFDH.getRecSetRd(querySQL)
		If m_rsDiv Is Nothing Then
			m_oCtrlErr.Import( m_oDbFDH.getObjErr() )
			Read = -1
			Exit Function
		End If
		If Not m_rsDiv.Eof Then
			' Vars
			getValues()
			Read = 1
		Else
			Read = 0
		End If

		m_rsDiv.Close()

	End Function


	' Write
	Public Function Write( dt, exist )

		' Verify Arguments
		If m_strGTAR = "" Or _
		   m_strRBAC = "" Or _
		   m_sSolCodi = "" Then
			m_oCtrlErr.Error = "Invalid Arguments."
			Write = -1
			Exit Function
		End If

		Dim querySQL, ret

		If exist Then
			querySQL =	"UPDATE AIRStatistics " & _
						" SET AIRStats_ANAC = " & m_nVarANAC & ", " & _
						"  AIRStats_30D_CLOSED = " & m_nVar30dClosed & ", " & _
						"  AIRStats_30D_DOCS = " & m_nVar30dDocs & ", " & _
						"  AIRStats_MAX_ANAC = " & m_nVarMaxANAC & ", " & _
						"  AIRStats_30D_ANAC = " & m_nVar30dANAC & ", " & _
						"  AIRStats_60D_ANAC = " & m_nVar60dANAC & ", " & _
						"  AIRStats_CLIENT = " & m_nVarClient & ", " & _
						"  AIRStats_DELAYED_CLIENT = " & m_nVarDelayClient & ", " & _
						"  AIRStats_MAX_CLIENT = " & m_nVarMaxClient & ", " & _
						"  AIRStats_DELIVERY = " & m_nVarDelivery & ", " & _
						"  AIRStats_7D_DELIVERY = " & m_nVar7dDelivery & ", " & _
						"  AIRStats_14D_DELIVERY = " & m_nVar14dDelivery & ", " & _
						"  AIRStats_MAX_DELIVERY = " & m_nVarMaxDelivery & ", " & _
						"  AIRStats_TIMESTAMP = '" & Now() & "'" & _
						" WHERE AIRStats_DATE=#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "# AND " & _
						"  AIRStats_SOLIC='" & m_sSolCodi & "' AND " & _
						"  AIRStats_GTAR='" & m_strGTAR & "' AND " & _
						"  AIRStats_RBAC='" & m_strRBAC & "'"
		Else
			querySQL =	"INSERT INTO AIRStatistics " & _
						" (AIRStats_DATE, AIRStats_SOLIC, AIRStats_GTAR, AIRStats_RBAC, AIRStats_ANAC, AIRStats_CLIENT, AIRStats_DELIVERY, " & _
						"  AIRStats_30D_ANAC, AIRStats_60D_ANAC, AIRStats_MAX_ANAC, AIRStats_DELAYED_CLIENT, AIRStats_MAX_CLIENT, AIRStats_7D_DELIVERY, " & _
						"  AIRStats_14D_DELIVERY, AIRStats_MAX_DELIVERY, AIRStats_30D_CLOSED, AIRStats_30D_DOCS, " & _
						"  AIRStats_TIMESTAMP) " & _
						" VALUES (#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "#, '" & _
						m_sSolCodi & "', '" & m_strGTAR & "', '" & m_strRBAC & "', " & m_nVarANAC & ", " & _
						m_nVarClient & ", " & m_nVarDelivery & ", " & _
						m_nVar30dANAC & ", " & m_nVar60dANAC & ", " & m_nVarMaxANAC & ", " & m_nVarDelayClient & ", " & _
						m_nVarMaxClient & ", " & m_nVar7dDelivery & ", " & m_nVar14dDelivery & ", " & m_nVarMaxDelivery & ", " & _
						m_nVar30dClosed & ", " & m_nVar30dDocs & ", '" & Now() & "');"
		End If
		ret = m_oDbFDH.Execute( querySQL )
		If ret Is Nothing then
			m_oCtrlErr.Import( m_oDbFDH.getObjErr() )
			Write = -1
			Exit Function
		End If

		Write = 1

	End Function


	' Close
	Public Sub Close()
		m_oDbFDH.Close()
	End Sub

	'  Get Error
	Public Property Get ErrorNumber()
		ErrorNumber = m_oCtrlErr.ErrorNumber
	End Property

	Public Property Get ErrorDescr()
		ErrorDescr = m_oCtrlErr.ErrorDescr
	End Property

	'  Get Error Object
	Public Function getObjErr()
		Set getObjErr = m_oCtrlErr
	End Function

	'  Print Error
	Public Sub Print
		m_oCtrlErr.Print()
	End Sub

End Class

%>
