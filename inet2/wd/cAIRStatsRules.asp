<%

'----------------------------------------------------------------
'
'	Class cAIRStatsRules
'
'	Date: 27/07/2014
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------

'----------------------------------------------------------------
'
'	Class cAIRStatsRules
'
Class cAIRStatsRules

	'Declarations
	Dim m_rVal(3,5,7)

	' Variables

	' Indicators/Results
	Private m_resDelayANAC		' Indica atraso da ANAC
	Private m_resDelayClient	' Indica atraso da Empresa
	Private m_resDelayDelivery	' Indica atraso de Distribuição

	'Class Initialization
	Private Sub Class_Initialize()
		m_resDelayANAC		= 0
		m_resDelayClient	= 0
		m_resDelayDelivery	= 0
		' default rules values
		Dim i, j, k
		For i=0 To 2
			For j=0 To 4
				For k=0 To 6
					m_rVal(i,j,k) = 0
				Next 
			Next 
		Next 
	End Sub
	Public Default Function construct( parameters )
		If UBound(parameters) = 105 Then
			Dim tag : tag = parameters(0)
			Dim i, j, k, ret, parse, key
			Dim l : l = 1
			' Default
			For i=0 To 2
				For j=0 To 4
					For k=0 To 6
						m_rVal(i,j,k) = CDbl(parameters(l))
						l = l + 1
					Next
				Next
			Next
			' DB Connection
			Dim oDbFDH : Set oDbFDH = (new cDBAccess)("FDH")
			Dim querySQL : querySQL =	"SELECT * FROM WDRules WHERE WDRules_Id = '" & tag & "'"
			Dim rsDiv : Set rsDiv = oDbFDH.getRecSetRd(querySQL)
			If Not rsDiv.Eof Then
				For i=0 To 2
					Select Case i
						Case 0
							ret = rsDiv("WDRules_Anac")
						Case 1
							ret = rsDiv("WDRules_Client")
						Case 2
							ret = rsDiv("WDRules_Delivery")
					End Select
					parse = Split(ret, ";")
					Dim p : p = 0
					For j=0 To 4
						For k=0 To 6
							If p < UBound( parse ) Then
								m_rVal(i,j,k) = CDbl(parse(p))
								p = p + 1
							End If
						Next
					Next
				Next
			Else
				' Salva valores
				Dim exec(3) : exec(0) = "": exec(1) = "": exec(2) = ""
				For i=0 To 2
					For j=0 To 4
						For k=0 To 6
							exec(i) = exec(i) & CStr(m_rVal(i,j,k)) & ";"
						Next
					Next
				Next
				querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery) VALUES ('" & tag & "', '" & exec(0) & "', '" & exec(1) & "', '" & exec(2) & "');"
				ret = oDbFDH.Execute( querySQL )
			End If
			rsDiv.Close()
			oDbFDH.Close()
		End If
        Set construct = Me
    End Function

	'Terminate Class
	Private Sub Class_Terminate()
	End Sub

	' Max Array
	Private Function MaxArray(A)
		Dim I
		MaxArray = A(LBound(A))
		For I = LBound(A) + 1 To UBound(A)
			If A(I) > MaxArray Or IsNull(MaxArray) Then MaxArray = A(I)
		Next
	End Function

	'  Get m_resDelayANAC
	Public Property Get resDelayANAC()
		resDelayANAC = m_resDelayANAC
	End Property

	'  Get m_resDelayClient
	Public Property Get resDelayClient()
		resDelayClient = m_resDelayClient
	End Property

	'  Get m_resDelayDelivery
	Public Property Get resDelayDelivery()
		resDelayDelivery = m_resDelayDelivery
	End Property

	'  Get resDelayMax
	Public Property Get resDelayMax()
		resDelayMax = MaxArray(Array(m_resDelayANAC,m_resDelayClient,m_resDelayDelivery))
	End Property

	Public Function getValues( i, j, k )
		getValues = m_rVal(i,j,k)
    End Function


	' Regras para cálculo dos indicadores
	Public Sub Rules( oAIRData )

		'--------------------------------------------------------------------------
		' m_resDelayANAC	- Indica atraso da ANAC na analise dos processos
		If oAIRData.ANAC = 0 Then
			m_resDelayANAC = RES_NONE

		' CRITICAL
		ElseIf oAIRData.ANAC > m_rVal(0,0,0) And _
			   ( oAIRData.ANACDelay > m_rVal(0,0,1) Or _
				 oAIRData.ANACMaxDays > m_rVal(0,0,2) Or _ 
				 ( oAIRData.ANAC > m_rVal(0,0,3) And _
				   oAIRData.ANAC60d/oAIRData.ANAC > m_rVal(0,0,4) ) Or _
				 ( oAIRData.ANAC > m_rVal(0,0,5) And _
				   oAIRData.ANAC30d/oAIRData.ANAC > m_rVal(0,0,6) ) ) Then
				m_resDelayANAC = RES_CRITICO

		' RES_RUIM
		ElseIf oAIRData.ANAC > m_rVal(0,1,0) And _
			   ( oAIRData.ANACDelay > m_rVal(0,1,1) Or _
				 oAIRData.ANACMaxDays > m_rVal(0,1,2) Or _ 
				 ( oAIRData.ANAC > m_rVal(0,1,3) And _
				   oAIRData.ANAC60d/oAIRData.ANAC > m_rVal(0,1,4) ) Or _
				 ( oAIRData.ANAC > m_rVal(0,1,5) And _
				   oAIRData.ANAC30d/oAIRData.ANAC > m_rVal(0,1,6) ) ) Then
				m_resDelayANAC = RES_RUIM

		' RES_ALERTA
		ElseIf oAIRData.ANAC > m_rVal(0,2,0) And _
			   ( oAIRData.ANACDelay > m_rVal(0,2,1) Or _
				 oAIRData.ANACMaxDays > m_rVal(0,2,2) Or _ 
				 ( oAIRData.ANAC > m_rVal(0,2,3) And _
				   oAIRData.ANAC60d/oAIRData.ANAC > m_rVal(0,2,4) ) Or _
				 ( oAIRData.ANAC > m_rVal(0,2,5) And _
				   oAIRData.ANAC30d/oAIRData.ANAC > m_rVal(0,2,6) ) ) Then
				m_resDelayANAC = RES_ALERTA

		' RES_BOM
		ElseIf oAIRData.ANAC > m_rVal(0,3,0) And _
			   ( oAIRData.ANACDelay > m_rVal(0,3,1) Or _
				 oAIRData.ANACMaxDays > m_rVal(0,3,2) Or _ 
				 ( oAIRData.ANAC > m_rVal(0,3,3) And _
				   oAIRData.ANAC60d/oAIRData.ANAC > m_rVal(0,3,4) ) Or _
				 ( oAIRData.ANAC > m_rVal(0,3,5) And _
				   oAIRData.ANAC30d/oAIRData.ANAC > m_rVal(0,3,6) ) ) Then
				m_resDelayANAC = RES_BOM

		' RES_OTIMO
		ElseIf oAIRData.ANAC > m_rVal(0,4,0) And _
			   ( oAIRData.ANACDelay > m_rVal(0,4,1) Or _
				 oAIRData.ANACMaxDays > m_rVal(0,4,2) Or _ 
				 ( oAIRData.ANAC > m_rVal(0,4,3) And _
				   oAIRData.ANAC60d/oAIRData.ANAC > m_rVal(0,4,4) ) Or _
				 ( oAIRData.ANAC > m_rVal(0,4,5) And _
				   oAIRData.ANAC30d/oAIRData.ANAC > m_rVal(0,4,6) ) ) Then
				m_resDelayANAC = RES_OTIMO

		' RES_NONE
		Else
			m_resDelayANAC = RES_NONE
		End If


		'--------------------------------------------------------------------------
		' m_resDelayClient	- Indica atraso da OM na resposta dos processos
		If oAIRData.Client = 0 Then
			m_resDelayClient = RES_NONE

		' CRITICAL
		ElseIf oAIRData.Client > m_rVal(1,0,0) And _
			   ( oAIRData.ClientDelay > m_rVal(1,0,1) Or _
				 oAIRData.ClientMaxDays > m_rVal(1,0,2) ) Then
				m_resDelayClient = RES_CRITICO

		' RES_RUIM
		ElseIf oAIRData.Client > m_rVal(1,1,0) And _
			   ( oAIRData.ClientDelay > m_rVal(1,1,1) Or _
				 oAIRData.ClientMaxDays > m_rVal(1,1,2) ) Then
				m_resDelayClient = RES_RUIM

		' RES_ALERTA
		ElseIf oAIRData.Client > m_rVal(1,2,0) And _
			   ( oAIRData.ClientDelay > m_rVal(1,2,1) Or _
				 oAIRData.ClientMaxDays > m_rVal(1,2,2) ) Then
				m_resDelayClient = RES_ALERTA

		' RES_BOM
		ElseIf oAIRData.Client > m_rVal(1,3,0) And _
			   ( oAIRData.ClientDelay > m_rVal(1,3,1) Or _
				 oAIRData.ClientMaxDays > m_rVal(1,3,2) ) Then
				m_resDelayClient = RES_BOM

		' RES_OTIMO
		ElseIf oAIRData.Client > m_rVal(1,4,0) And _
			   ( oAIRData.ClientDelay > m_rVal(1,4,1) Or _
				 oAIRData.ClientMaxDays > m_rVal(1,4,2) ) Then
				m_resDelayClient = RES_OTIMO

		' RES_NONE
		Else
			m_resDelayClient = RES_NONE
		End If


		'--------------------------------------------------------------------------
		' m_resDelayDelivery	- Indica atraso na distribuição de processos
		If oAIRData.Delivery = 0 Then
			m_resDelayDelivery = RES_NONE

		' CRITICAL
		ElseIf oAIRData.Delivery > m_rVal(2,0,0) And _
			   ( oAIRData.DeliveryDelay > m_rVal(2,0,1) Or _
				 oAIRData.DeliveryMaxDays > m_rVal(2,0,2) Or _ 
				 ( oAIRData.Delivery > m_rVal(2,0,3) And _
				   oAIRData.Delivery14d/oAIRData.Delivery > m_rVal(2,0,4) ) Or _
				 ( oAIRData.Delivery > m_rVal(2,0,5) And _
				   oAIRData.Delivery7d/oAIRData.Delivery > m_rVal(2,0,6) ) ) Then
				m_resDelayDelivery = RES_CRITICO

		' RES_RUIM
		ElseIf oAIRData.Delivery > m_rVal(2,1,0) And _
			   ( oAIRData.DeliveryDelay > m_rVal(2,1,1) Or _
				 oAIRData.DeliveryMaxDays > m_rVal(2,1,2) Or _ 
				 ( oAIRData.Delivery > m_rVal(2,1,3) And _
				   oAIRData.Delivery14d/oAIRData.Delivery > m_rVal(2,1,4) ) Or _
				 ( oAIRData.Delivery > m_rVal(2,1,5) And _
				   oAIRData.Delivery7d/oAIRData.Delivery > m_rVal(2,1,6) ) ) Then
				m_resDelayDelivery = RES_RUIM

		' RES_ALERTA
		ElseIf oAIRData.Delivery > m_rVal(2,2,0) And _
			   ( oAIRData.DeliveryDelay > m_rVal(2,2,1) Or _
				 oAIRData.DeliveryMaxDays > m_rVal(2,2,2) Or _ 
				 ( oAIRData.Delivery > m_rVal(2,2,3) And _
				   oAIRData.Delivery14d/oAIRData.Delivery > m_rVal(2,2,4) ) Or _
				 ( oAIRData.Delivery > m_rVal(2,2,5) And _
				   oAIRData.Delivery7d/oAIRData.Delivery > m_rVal(2,2,6) ) ) Then
				m_resDelayDelivery = RES_ALERTA

		' RES_BOM
		ElseIf oAIRData.Delivery > m_rVal(2,3,0) And _
			   ( oAIRData.DeliveryDelay > m_rVal(2,3,1) Or _
				 oAIRData.DeliveryMaxDays > m_rVal(2,3,2) Or _ 
				 ( oAIRData.Delivery > m_rVal(2,3,3) And _
				   oAIRData.Delivery14d/oAIRData.Delivery > m_rVal(2,3,4) ) Or _
				 ( oAIRData.Delivery > m_rVal(2,3,5) And _
				   oAIRData.Delivery7d/oAIRData.Delivery > m_rVal(2,3,6) ) ) Then
				m_resDelayDelivery = RES_BOM

		' RES_OTIMO
		ElseIf oAIRData.Delivery > m_rVal(2,4,0) And _
			   ( oAIRData.DeliveryDelay > m_rVal(2,4,1) Or _
				 oAIRData.DeliveryMaxDays > m_rVal(2,4,2) Or _ 
				 ( oAIRData.Delivery > m_rVal(2,4,3) And _
				   oAIRData.Delivery14d/oAIRData.Delivery > m_rVal(2,4,4) ) Or _
				 ( oAIRData.Delivery > m_rVal(2,4,5) And _
				   oAIRData.Delivery7d/oAIRData.Delivery > m_rVal(2,4,6) ) ) Then
				m_resDelayDelivery = RES_OTIMO

		' RES_NONE
		Else
			m_resDelayDelivery = RES_NONE
		End If

	End Sub

End Class

%>
