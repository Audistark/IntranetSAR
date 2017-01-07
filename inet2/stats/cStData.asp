<%

'----------------------------------------------------------------
'
'	Class cStData
'
'	Date: 03/05/2015
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------

'----------------------------------------------------------------
'
'	Class cStData
'
'	Guarda estatísticas relacionada a cada grupo estatístico cadastrado
'

Class cStData

	'Declarations
	Private m_oCtrlErr			' Error Object
	Private m_oDbFDH			' DB Object
	Private m_rsDiv				' rsDiv

	' Variables
	Private m_nDate				' Date
	Private m_sGTAR				' GTAR
	Private m_sRBAC				' RBAC
	Private m_sType				' Type 'S' or 'T'
	Private m_sCodi				' Code of Request

	Private m_nVarANAC			' Processos com a ANAC (em análise com a ANAC)
	Private m_nVarANACGoalDays	' Tempo meta da ANAC (Cada Grupo estatístico só pode ter uma Meta)
	Private m_nVarANACDelay		' Acima da meta com ANAC
	Private m_nVarANACMax		' Maior atraso em dias com a ANAC
	Private m_nVarANACDays		' Tempo em dias com ANAC

	Private m_nVar30dClosed		' Fechados nos últimos 30 dias (32 dias)

Private m_nVarItrClosed		' Taxa média de Iterações por processo fechados últimos 30 dias

	Private m_nVar30dDocs		' Documentos emitidos nos últimos 30 dias (32 dias)

	Private m_nVarClient		' Processos com a empresa
	Private m_nVarDelayClient	' Atrasados com Empresa (> que pendencia+30 ou 90+30)
	Private m_nVarMaxClient		' Maior atraso em dias com Empresa

	Private m_sTopTen(10,2)		' Solicitações mais atrasadas

	Private m_tStamp			' Timestamp

	'Class Initialization
	Private Sub Class_Initialize()
		Set m_oCtrlErr	= new cCtrlErr
	End Sub
	Public Default Function construct()
		m_nDate				= Date()
		m_sGTAR				= ""
		m_sRBAC				= ""
		m_sType				= ""
		m_sCodi				= ""
		Call CleanValues()
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

	'  Get Codi
	Public Property Get Codi()
		Codi = m_sCodi
	End Property
	Public Property Let Codi( value )
		m_sCodi = CInt(value)
		m_sCodi = String( 3 - Len(m_sCodi), "0" ) & m_sCodi
	End Property

	'  Get Type
	Public Property Get sType()
		sType = m_sType
	End Property
	Public Property Let sType( value )
		m_sType = value
	End Property

	'  Get GTAR
	Public Property Get GTAR()
		GTAR = m_sGTAR
	End Property
	Public Property Let GTAR( value )
		m_sGTAR = UCase(value)
	End Property

	'  Get RBAC
	Public Property Get RBAC()
		RBAC = m_sRBAC
	End Property
	Public Property Let RBAC( value )
		m_sRBAC = value
	End Property

	'  Get nVarANAC			' Processos com a ANAC (em análise com a ANAC)
	Public Property Get ANAC()
		ANAC = m_nVarANAC
	End Property
	Public Property Let ANAC( value )
		m_nVarANAC = value
	End Property

	'  Get nVarGoalDays			' Tempo meta da ANAC
	Public Property Get Goal()
		Goal = m_nVarANACGoalDays
	End Property
	Public Property Let Goal( value )
		m_nVarANACGoalDays = value
	End Property

	'  Get nVarANACDelay		' Acima da Meta com ANAC
	Public Property Get ANACDelay()
		ANACDelay = m_nVarANACDelay
	End Property
	Public Property Let ANACDelay( value )
		m_nVarANACDelay = value
	End Property

	'  Get nVarANACMax			' Acumulado de dias com ANAC
	Public Property Get ANACMaxDays()
		ANACMaxDays = m_nVarANACMax
	End Property
	Public Property Let ANACMaxDays( value )
		m_nVarANACMax = value
	End Property

	'  Get nVarANACDays		' Tempo em dias com ANAC
	Public Property Get ANACDays()
		ANACDays = m_nVarANACDays
	End Property
	Public Property Let ANACDays( value )
		m_nVarANACDays = value
	End Property

	'  Tempo médio com ANAC
	Public Property Get ANACAvg()
		If m_nVarANAC > 0 Then
			ANACAvg = CInt( m_nVarANACDays / m_nVarANAC )
		Else
			ANACAvg = 0
		End If
	End Property

	'  Get nVar30dClosed		' Fechados nos últimos 30 dias (32 dias)
	Public Property Get Closed30d()
		Closed30d = m_nVar30dClosed
	End Property
	Public Property Let Closed30d( value )
		m_nVar30dClosed = value
	End Property

'  Get nVarItrClosed		' Taxa média de Iterações por processo fechados últimos 30 dias
Public Property Get ItrClosedANAC()
	ItrClosedANAC = m_nVarItrClosed
End Property
Public Property Let ItrClosed( value )
	m_nVarItrClosed = value
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

	'  Get nVarSum				' Total de Processos (ANAC + Empresa)
	Public Property Get Sum()
		Sum = m_nVarANAC + m_nVarClient
	End Property

	'  Get m_tStamp				' Timestamp
	Public Property Get tStamp()
		tStamp = m_tStamp
	End Property

	'  Get TopTenProcSolic
	Public Function TopTenProcSolic( i )
		TopTenProcSolic = m_sTopTen(i,0)
	End Function
	
	'  Get TopTenDelay
	Public Function TopTenDelay( i )
		TopTenDelay = m_sTopTen(i,1)
	End Function


	' Get
	Private Sub getValues()
		m_nDate				= m_rsDiv( "Stats_DATE" )
		m_nVarANAC			= m_rsDiv( "Stats_ANAC" )
		m_nVarANACGoalDays	= m_rsDiv( "Stats_GOAL" )
		m_nVarANACDelay		= m_rsDiv( "Stats_DELAY_ANAC" )
		m_nVarANACMax		= m_rsDiv( "Stats_MAX_ANAC" )
		m_nVarANACDays		= m_rsDiv( "Stats_DAYS_ANAC" )
		m_nVar30dClosed		= m_rsDiv( "Stats_30D_CLOSED" )
		m_nVarItrClosed		= m_rsDiv( "Stats_ITR_CLOSED" )
		m_nVar30dDocs		= m_rsDiv( "Stats_30D_DOCS" )
		m_nVarClient		= m_rsDiv( "Stats_CLIENT" )
		m_nVarDelayClient	= m_rsDiv( "Stats_DELAYED_CLIENT" )
		m_nVarMaxClient		= m_rsDiv( "Stats_MAX_CLIENT" )
		Dim i, pos, sTopTen, sParse
		sTopTen = m_rsDiv( "Stats_TOPTEN" )
		If sTopTen <> "" And Not IsNull(sTopTen) Then
			sParse = split(sTopTen,";")
			For i=0 To UBound( sParse )
				pos = InStr(sParse(i),",")
				If pos > 0 Then
					m_sTopTen(i,0) = Left(sParse(i),pos-1)
					m_sTopTen(i,1) = CINt(Right(sParse(i),Len(sParse(i))-pos))
				End If
			Next
		End If
		m_tStamp			= m_rsDiv( "Stats_TIMESTAMP" )
	End Sub


	Public Sub CleanValues()
		m_nVarANAC			= 0
		m_nVarANACDelay		= 0
		m_nVarANACMax		= 0
		m_nVarANACDays		= 0
		m_nVar30dClosed		= 0
m_nVarItrClosed		= 0
		m_nVar30dDocs		= 0
		m_nVarClient		= 0
		m_nVarDelayClient	= 0
		m_nVarMaxClient		= 0
		Dim i
		For i=0 To 9
			m_sTopTen(i,0) = ""
			m_sTopTen(i,1) = 0
		Next
	End Sub


	' Import
	Public Sub Import( object )
		m_nDate				= object.dtDate
		m_sGTAR				= object.GTAR
		m_sRBAC				= object.RBAC
		m_sType				= object.sType
		m_sCodi				= object.Codi
		m_nVarANAC			= object.ANAC
		m_nVarANACGoalDays	= object.Goal
		m_nVarANACDelay		= object.ANACDelay
		m_nVarANACMax		= object.ANACMaxDays
		m_nVarANACDays		= object.ANACDays
		m_nVar30dClosed		= object.Closed30d
m_nVarItrClosed		= object.ItrClosedANAC
		m_nVar30dDocs		= object.Docs30d
		m_nVarClient		= object.Client
		m_nVarDelayClient	= object.ClientDelay
		m_nVarMaxClient		= object.ClientMaxDays
		Dim i
		For i=0 To 9
			m_sTopTen(i,0) = object.TopTenProcSolic(i)
			m_sTopTen(i,1) = object.TopTenDelay(i)
		Next
		m_tStamp			= object.tStamp
	End Sub


	' Add
	Public Sub Add( object )
		m_nVarANAC			= m_nVarANAC + object.ANAC
		m_nVarANACDelay		= m_nVarANACDelay + object.ANACDelay
		m_nVarANACDays		= m_nVarANACDays + object.ANACDays
		If object.ANACMaxDays > m_nVarANACMax Then m_nVarANACMax = object.ANACMaxDays
		m_nVar30dClosed		= m_nVar30dClosed + object.Closed30d
		m_nVar30dDocs		= m_nVar30dDocs + object.Docs30d
m_nVarItrClosed		= m_nVarItrClosed + object.ItrClosedANAC
		m_nVarClient		= m_nVarClient + object.Client
		m_nVarDelayClient	= m_nVarDelayClient + object.ClientDelay
		If object.ClientMaxDays > m_nVarMaxClient Then m_nVarMaxClient = object.ClientMaxDays
		Dim i
		For i=0 To 9
			Call TopTen( object.TopTenProcSolic(i), object.TopTenDelay(i) )
		Next
	End Sub


	Public Function TopTen( ProcSolic, Delay )

		If ProcSolic = "" Or Delay = 0 Then
			TopTen = 0
			Exit Function
		End If

		Dim i, j
		Dim nLastDelay : nLastDelay = 9999
		Dim iLastDelay : iLastDelay = 0
		Dim found : found = False
		For i=0 To 9
			If m_sTopTen(i,0) = ProcSolic Then
				If m_sTopTen(i,1) < Delay Then
					m_sTopTen(i,1) = Delay
					found = True
					Exit For
				End If
			ElseIf m_sTopTen(i,0) = "" Then
				iLastDelay = i
				Exit For
			ElseIf m_sTopTen(i,1) < nLastDelay Then
				nLastDelay = m_sTopTen(i,1)
				iLastDelay = i
			End If
		Next
		If found = False Then
			If iLastDelay < 0 Then
				TopTen = 0
				Exit Function
			End If
			m_sTopTen(iLastDelay,0) = ProcSolic
			m_sTopTen(iLastDelay,1) = Delay
		End If
		' sort
		For i=0 To 9
			For j=i+1 To 9
				If m_sTopTen(j,0) = "" Then
					Exit For
				ElseIf m_sTopTen(j,1) > m_sTopTen(i,1) Then
					ProcSolic = m_sTopTen(j,0)
					Delay = m_sTopTen(j,1)
					m_sTopTen(j,0) = m_sTopTen(i,0)
					m_sTopTen(j,1) = m_sTopTen(i,1)
					m_sTopTen(i,0) = ProcSolic
					m_sTopTen(i,1) = Delay
				End If
			Next
		Next

		TopTen = 1

	End Function


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
		querySQL =	"SELECT * FROM AIRStats " &_
					" WHERE Stats_DATE=#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "# AND " & _
					"  Stats_TYPE='" & m_sType & "' AND Stats_CODI='" & m_sCodi & "' AND " & _
					"  Stats_GTAR='" & m_sGTAR & "' AND Stats_RBAC='" & m_sRBAC & "'"
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
		Dim tStamp : tStamp = rsDiv( "Stats_TIMESTAMP" )
		Dim hLast : hLast = DateDiff("h", tStamp, Now())
		If hLast > 3 Then
			Exists = 0 ' not found - to force recalculation
		Else
			Exists = 1 ' already exists
		End If

		m_rsDiv.Close()

	End Function


	Public Function FetchStart( direction )

		' Verify Arguments
		If m_sGTAR = "" Or _
		    m_sRBAC = "" Or _
		     m_sType = "" Or _
		      m_sCodi = "" Then
			oCtrlErr.Error = "Invalid Arguments."
			FetchStart = -1
			Exit Function
		End If

		'-----------------------------------------
		' Verifica se já foi criado algum registro
		Dim querySQL
		querySQL =	"SELECT * FROM AIRStats " &_
					" WHERE " & _
					"  Stats_TYPE='" & m_sType & "' AND Stats_CODI='" & m_sCodi & "' AND " & _
					"  Stats_GTAR='" & m_sGTAR & "' AND Stats_RBAC='" & m_sRBAC & "'" & _
					" ORDER BY Stats_DATE"
		If direction < 0 Then ' reverse
			querySQL = querySQL & " DESC"
		End If
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
		On Error Resume Next
		m_rsDiv.MoveNext
		If Err.Number <> 0 Then
			FetchNext = -1 ' Closed
			Exit Function
		End If
		On Error GoTo 0
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
		If m_sGTAR = "" Or _
		    m_sRBAC = "" Or _
		     m_sType = "" Or _
		      m_sCodi = "" Then
			m_oCtrlErr.Error = "Invalid Arguments."
			Close()
			Read = -1
			Exit Function
		End If

		'--------------------------------------------------------------
		' Verifica se já foi criado algum registro para a data desejada
		Dim querySQL
		querySQL =	"SELECT * FROM AIRStats " &_
					" WHERE Stats_DATE=#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "# AND " & _
					"  Stats_TYPE='" & m_sType & "' AND Stats_CODI='" & m_sCodi & "' AND " & _
					"  Stats_GTAR='" & m_sGTAR & "' AND Stats_RBAC='" & m_sRBAC & "'"
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
		If m_sGTAR = "" Or _
		    m_sRBAC = "" Or _
		     m_sType = "" Or _
		      m_sCodi = "" Then
			m_oCtrlErr.Error = "Invalid Arguments."
			Write = -1
			Exit Function
		End If

		Dim querySQL, ret

		Dim i, sTopTen : sTopTen = ""
		For i=0 To 9
			If m_sTopTen(i,0) <> "" And m_sTopTen(i,1) > 0 Then
				sTopTen = sTopTen & m_sTopTen(i,0) & "," & m_sTopTen(i,1) & ";"
			Else
				Exit For
			End If
		Next

		'Dim dayStamp : dayStamp = Day(Date()) : If Len(dayStamp) < 2 Then dayStamp = "0" & dayStamp
		'Dim monthStamp : monthStamp = Month(Date()) : If Len(monthStamp) < 2 Then monthStamp = "0" & monthStamp
		'Dim yearStamp : yearStamp = Year(Date())
		'Dim hourStamp : hourStamp = Hour(Time()) : If Len(hourStamp) < 2 Then hourStamp = "0" & hourStamp
		'Dim minuteStamp : minuteStamp = Minute(Time()) : If Len(minuteStamp) < 2 Then minuteStamp = "0" & minuteStamp
		'Dim secondStamp : secondStamp = Second(Time()) : If Len(secondStamp) < 2 Then secondStamp = "0" & secondStamp
		'Dim sTStamp : sTStamp = dayStamp & "/" & monthStamp & "/" & yearStamp & " " & hourStamp & ":" & minuteStamp & ":" & minuteStamp

		'Set the server locale
		Session.LCID = 1046 'BRASIL - Formato de data brasileiro

		If exist Then
			querySQL =	"UPDATE AIRStats " & _
						" SET Stats_ANAC = " & m_nVarANAC & ", " & _
						"  Stats_GOAL = " & m_nVarANACGoalDays & ", " & _
						"  Stats_DELAY_ANAC = " & m_nVarANACDelay & ", " & _
						"  Stats_MAX_ANAC = " & m_nVarANACMax & ", " & _
						"  Stats_DAYS_ANAC = " & m_nVarANACDays & ", " & _
						"  Stats_30D_CLOSED = " & m_nVar30dClosed & ", " & _
						"  Stats_ITR_CLOSED = " & m_nVarItrClosed & ", " & _
						"  Stats_30D_DOCS = " & m_nVar30dDocs & ", " & _
						"  Stats_CLIENT = " & m_nVarClient & ", " & _
						"  Stats_DELAYED_CLIENT = " & m_nVarDelayClient & ", " & _
						"  Stats_MAX_CLIENT = " & m_nVarMaxClient & ", " & _
						"  Stats_TOPTEN = '" & sTopTen & "', " & _
						"  Stats_TIMESTAMP = '" & Now() & "'" & _
						" WHERE Stats_DATE=#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "# AND " & _
						"  Stats_TYPE='" & m_sType & "' AND Stats_CODI='" & m_sCodi & "' AND " & _
						"  Stats_GTAR='" & m_sGTAR & "' AND Stats_RBAC='" & m_sRBAC & "'"
		Else
			querySQL =	"INSERT INTO AIRStats " & _
						" (Stats_DATE, Stats_TYPE, Stats_CODI, Stats_GTAR, Stats_RBAC, " & _
						"  Stats_ANAC, Stats_GOAL, Stats_DELAY_ANAC, Stats_MAX_ANAC, Stats_DAYS_ANAC, " & _
						"  Stats_30D_CLOSED, Stats_ITR_CLOSED, Stats_30D_DOCS, " & _
						"  Stats_CLIENT, Stats_DELAYED_CLIENT, Stats_MAX_CLIENT, Stats_TOPTEN, " & _
						"  Stats_TIMESTAMP) " & _
						" VALUES (#" & Month(dt) & "/" & Day(dt) & "/" & Year(dt) & "#, '" & _
						m_sType & "', '" & m_sCodi & "', '" & m_sGTAR & "', '" & m_sRBAC & "', " & _
						m_nVarANAC & ", " & m_nVarANACGoalDays & ", " & m_nVarANACDelay & ", " & m_nVarANACMax & ", " & _
						m_nVarANACDays & ", " & m_nVar30dClosed & ", " & m_nVarItrClosed & ", " & _
						m_nVar30dDocs & ", " & m_nVarClient & ", " & m_nVarDelayClient & ", " & _
						m_nVarMaxClient & ", '" & sTopTen & "', '" & Now() & "');"

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
