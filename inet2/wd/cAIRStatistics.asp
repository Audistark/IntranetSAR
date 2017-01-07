<%

'----------------------------------------------------------------
'
'	Class cAIRStatistics
'
'	Date: 27/07/2014
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------
Const TST_145CODI_DELIVERY = "027"	' Em Distribuição
Const TST_135CODI_DELIVERY = "014"	' Em Distribuição


'----------------------------------------------------------------
'
'	Class cAIRStatistics
'
Class cAIRStatistics

	' Declarations
	Private m_oCtrlErr			' Error Object
	Private m_oAIRData			' Data Object

	' Class Initialization
	Private Sub Class_Initialize()
		Set m_oCtrlErr	= new cCtrlErr
		Set m_oAIRData	= new cAIRStatsData
	End Sub
	Public Default Function construct( parameters )
		' "145", "GTAR-DF", "S001", "S004", "S012", "T005", "T017"
		Dim ret : ret = UBound(parameters)
		If ret < 1 Then
			m_oCtrlErr.Error = "Invalid Arguments."
		ElseIf ret = 1 Then
			m_oAIRData.RBAC = parameters(0)
			m_oAIRData.GTAR	= parameters(1)
			m_oAIRData.SCodi= "000"
		Else
			m_oAIRData.RBAC = parameters(0)
			m_oAIRData.GTAR	= parameters(1)
			m_oAIRData.SCodi= parameters(2)
		End If
        Set construct = Me
    End Function

	' Terminate Class
	Private Sub Class_Terminate()
		Set m_oCtrlErr = Nothing
	End Sub

	'  Get oData
	Public Property Get oData()
		Set oData = m_oAIRData
	End Property

	' Open Connection
	Public Function GetValues( dt )

		If m_oAIRData.Open() < 0 then
			m_oCtrlErr.Import( m_oAIRData.getObjErr() )
			GetValues = m_oCtrlErr.ErrorNumber
			Exit Function
		End If

		Dim tStamp : tStamp = Now()
		Dim ret : ret = -1
		Dim bExist : bExist = False
		Dim docman_ant : docman_ant = ""

		'--------------------------------------------------------------
		' Verifica se já foi criado algum registro para a data desejada
		ret = m_oAIRData.Read(dt)
		If ret < 0 Then
			m_oCtrlErr.Import( m_oAIRData.getObjErr() )
			GetValues = m_oCtrlErr.ErrorNumber
			Exit Function
		End If
		If ret > 0 Then

			' Yes it exists
			bExist = True

			' É hoje?
			If dt <> Date() Then
				m_oAIRData.Close()
				GetValues = 0
				Exit Function
			End If

			' Previne overrun (3h)
			If ( DateDiff("h", m_oAIRData.tStamp, Now()) < 3 ) Then
				m_oAIRData.Close()
				GetValues = 0
				Exit Function
			End If

		End If

		m_oAIRData.CleanValues()

		Dim oDbFDH, rsDiv, querySQL
		Set oDbFDH = m_oAIRData.oDbFDH

		'
		' Busca valores
		'
		querySQL =	"SELECT O.ORG_CODI, O.ORGP_CODI, O.ORG_NABREV, P.P" & m_oAIRData.RBAC & "_CODI, P.P" & m_oAIRData.RBAC & "_DOCMAN, B.B" & m_oAIRData.RBAC & "_CODI, " & _
					"       S.S" & m_oAIRData.RBAC & "_DTSTAT, S.S" & m_oAIRData.RBAC & "_CODI, S.TST_CODI, TS.TSOL_DESCR, TSt.TST_DESCR, " & _
					"       TSt.TST_ORIGEM, TSt.TST_PROSSEGUE AS SOLIC_NCLOSED, T.P" & m_oAIRData.RBAC & "_S" & m_oAIRData.RBAC & ", T.T" & m_oAIRData.RBAC & "_CODI, " & _
					"       TSDiv.SDIV_SIGLA, Pes.PES_NGUERRA, " & _
					"       T.T" & m_oAIRData.RBAC & "_DATA, T.T" & m_oAIRData.RBAC & "_DTSTAT, T.T" & m_oAIRData.RBAC & "_DTPEND, TT.TSK_DESCR, " & _
					"       TSt_1.TST_PROSSEGUE AS TASK_NCLOSED, TSt_1.TST_CONTA, TSt_1.TST_REQDT, " & _
					"       D.D" & m_oAIRData.RBAC & "_ORIGEM, D.D" & m_oAIRData.RBAC & "_DATA, D.D" & m_oAIRData.RBAC & "_DOCMAN " & _
					"    FROM ( ( ( ( ( ( ( ( ( ( Organizacao AS O INNER JOIN A" & m_oAIRData.RBAC & "_Bases AS B ON O.ORG_CODI = B.ORG_CODI ) " & _
					"                           INNER JOIN A" & m_oAIRData.RBAC & "_Processos AS P ON B.B" & m_oAIRData.RBAC & "_CODI = P.B" & m_oAIRData.RBAC & "_CODI ) " & _
					"                         INNER JOIN A" & m_oAIRData.RBAC & "_Documentos AS D ON D.P" & m_oAIRData.RBAC & "_CODI = P.P" & m_oAIRData.RBAC & "_CODI ) " & _
					"                       INNER JOIN A" & m_oAIRData.RBAC & "_Solicitacoes AS S ON P.P" & m_oAIRData.RBAC & "_CODI = S.P" & m_oAIRData.RBAC & "_CODI ) " & _
					"                     INNER JOIN A" & m_oAIRData.RBAC & "_TabSolic AS TS ON S.TSOL_CODI = TS.TSOL_CODI ) " & _
					"                   LEFT JOIN A" & m_oAIRData.RBAC & "_Tarefas AS T ON S.P" & m_oAIRData.RBAC & "_S" & m_oAIRData.RBAC & " = T.P" & m_oAIRData.RBAC & "_S" & m_oAIRData.RBAC & " ) " & _
					"                 LEFT JOIN A" & m_oAIRData.RBAC & "_TabStatus AS TSt_1 ON T.TST_CODI = TSt_1.TST_CODI ) " & _
					"               LEFT JOIN A" & m_oAIRData.RBAC & "_TabTarefa AS TT ON T.TSK_CODI = TT.TSK_CODI ) " & _
					"             INNER JOIN Pessoal AS Pes ON S.PES_CODI = Pes.PES_CODI ) " & _
					"           INNER JOIN A" & m_oAIRData.RBAC & "_TabStatus AS TSt ON S.TST_CODI = TSt.TST_CODI ) " & _
					"         INNER JOIN Tab_Subdivisao AS TSDiv ON Pes.SDIV_CODI = TSDiv.SDIV_CODI " & _
					"    WHERE TS.TSOL_CODI='" & m_oAIRData.SCodi & "' AND TSDiv.SDIV_SIGLA='" & m_oAIRData.GTAR & "' AND " & _
					"         ( TSt.TST_PROSSEGUE='S' OR DateDiff( 'd', S.S" & m_oAIRData.RBAC & "_DTSTAT, Now) < 32 ) " & _
					"    ORDER BY P.P" & m_oAIRData.RBAC & "_CODI, S.S" & m_oAIRData.RBAC & "_CODI, T.T" & m_oAIRData.RBAC & "_CODI, D.D" & m_oAIRData.RBAC & "_DATA" 
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)

		If rsDiv Is Nothing Then
			m_oCtrlErr.Import( oDbFDH.getObjErr() )
			m_oAIRData.Close()
			GetValues = m_oCtrlErr.ErrorNumber
			Exit Function
		End If

		If Not rsDiv.Eof Then

			Dim FlagProc : FlagProc = True
			Dim FlagEOReg : FlagEOReg = False
			Dim PCodi : PCodi = rsDiv( "P" & m_oAIRData.RBAC & "_CODI" )
			Dim SCodi : SCodi = rsDiv( "S" & m_oAIRData.RBAC & "_CODI" )
			Dim StOrigin : StOrigin = rsDiv( "TST_ORIGEM" ) ' Origem Solicitação 'A' ou 'E'
			Dim StSolic : StSolic = rsDiv("TST_CODI")
			Dim DtSolic : DtSolic = rsDiv("S" & m_oAIRData.RBAC & "_DTSTAT")
			Dim ProcSolic : ProcSolic = PCodi & SCodi
			Dim OrgCodi : OrgCodi = rsDiv( "ORG_CODI" )
			Dim TskIni : TskIni = rsDiv("T" & m_oAIRData.RBAC & "_CODI")
			Dim DataPend : DataPend   = ""
			Dim DataAudit : DataAudit = ""

			While FlagProc

				Dim DtTaskExec : DtTaskExec = rsDiv("T" & m_oAIRData.RBAC & "_DATA")
				Dim DtTaskPend : DtTaskPend = rsDiv("T" & m_oAIRData.RBAC & "_DTPEND")

				Dim bSolicProssegue : bSolicProssegue = False
				If rsDiv("SOLIC_NCLOSED") = "S" Then ' Não concluído
					bSolicProssegue = True
				End If

				Dim bStProssegue : bStProssegue = False
				If rsDiv("TASK_NCLOSED") = "S" Then ' Não concluído
					bStProssegue = True
				End If

				Dim bContaTempo : bContaTempo = False
				If rsDiv("TST_CONTA") = "S" Then ' sim
					bContaTempo = True
				End If

				Dim bTaskDate : bTaskDate = False
				If rsDiv("TST_REQDT") = "S" Then
					bTaskDate = True
				End If

				' Verifica emissão de documentos pela ANAC nos últimos 30 dias
				If rsDiv("D" & m_oAIRData.RBAC & "_ORIGEM") = "A" Then
					If rsDiv("T" & m_oAIRData.RBAC & "_CODI") = TskIni Then
						Dim docman : docman = rsDiv("D" & m_oAIRData.RBAC & "_DOCMAN")
						If DateDiff("d", rsDiv("D" & m_oAIRData.RBAC & "_DATA"), Date()) < 32 And _
						   docman <> docman_ant Then
							m_oAIRData.Docs30d = m_oAIRData.Docs30d + 1
							docman_ant = docman
						End If
					End If
				End If

				' procura por data de pendência mais próxima em tarefas não concluidas..
				' e atualiza a data da solicitação
				If bStProssegue = True And bContaTempo = True Then ' Não concluído e conta tempo

					' Se tem auditoria agendada no futuro
					If DtTaskExec <> "" And bTaskDate = True And _
					   ( DataAudit = "" Or DtTaskExec > DataAudit ) Then
						If ( DtTaskExec + 10 ) >= Date() Then
							DataAudit = DtTaskExec
						End If
					End If

					' procura pela data de pendência mais nova..
					If StOrigin = "E" And _
					   DtTaskPend <> "" And _
					   ( DataPend = "" Or DtTaskPend < DataPend ) Then
						DataPend = DtTaskPend
					End If

				End If

				' get next
				rsDiv.MoveNext

				FlagEOReg = False ' default

				If rsDiv.Eof Then
					FlagProc = False
					FlagEOReg = True
				Else
					Dim OrgCodiNew : OrgCodiNew = rsDiv( "ORG_CODI" )
					Dim PCodiNew : PCodiNew = rsDiv( "P" & m_oAIRData.RBAC & "_CODI" )
					Dim SCodiNew : SCodiNew = rsDiv( "S" & m_oAIRData.RBAC & "_CODI" )
					Dim ProcSolicNew : ProcSolicNew = PCodiNew & SCodiNew
					If OrgCodi <> OrgCodiNew Or _
					   ProcSolic <> ProcSolicNew Then
						FlagEOReg = True
					End If
				End If

				' Finaliza Cálculos
				If FlagEOReg = True Then

					' Data Solic devido a Auditoria
					If StOrigin = "A" And _
						DataAudit <> "" And _
						DataAudit > DtSolic Then
						DtSolic = DataAudit
					End If

					' Verifica solicitações fechadas nos últimos 30 dias
					If bSolicProssegue = False Then
						m_oAIRData.Closed30d = m_oAIRData.Closed30d + 1
					Else

						Dim Delay : Delay = DateDiff("d", DtSolic, Date())

						If DataPend <> "" Then
							If DateDiff("d", DataPend, Date()) > 15 Then
								If DataPend > DtSolic Then
									Delay = DateDiff("d", DataPend, Date())
								End If
							Else
								Delay = 0
							End If
						End If

						' Delivery (só 135 e 145)
						If ( m_oAIRData.RBAC = "135" And StSolic = TST_135CODI_DELIVERY ) Or _
							( m_oAIRData.RBAC = "145" And StSolic = TST_145CODI_DELIVERY ) Then
							m_oAIRData.Delivery = m_oAIRData.Delivery + 1
							If Delay > m_oAIRData.DeliveryMaxDays Then m_oAIRData.DeliveryMaxDays = Delay
							If Delay > 7 Then
								m_oAIRData.Delivery7d = m_oAIRData.Delivery7d + 1
							End If
							If Delay > 14 Then
								m_oAIRData.Delivery14d = m_oAIRData.Delivery14d + 1
							End If
						' ANAC
						ElseIf StOrigin = "A" Then
							m_oAIRData.ANAC = m_oAIRData.ANAC + 1
							If Delay > m_oAIRData.ANACMaxDays Then m_oAIRData.ANACMaxDays = Delay
							If Delay > 30 Then
								m_oAIRData.ANAC30d = m_oAIRData.ANAC30d + 1
							End If
							If Delay > 60 Then
								m_oAIRData.ANAC60d = m_oAIRData.ANAC60d + 1
							End If
						' Empresa
						Else
							m_oAIRData.Client = m_oAIRData.Client + 1
							If DataPend = "" Then
								If Delay > (90+30) Then
									m_oAIRData.ClientDelay = m_oAIRData.ClientDelay + 1
								End If
								If ( Delay - 90 ) > m_oAIRData.ClientMaxDays Then
									m_oAIRData.ClientMaxDays = Delay - 90
								End If							
							Else
								If Delay > 30 Then		' Dá uma folga no prazo da pendência
									m_oAIRData.ClientDelay = m_oAIRData.ClientDelay + 1
								End If
								If Delay > m_oAIRData.ClientMaxDays Then
									m_oAIRData.ClientMaxDays = Delay
								End If
							End If
						End If

					End If

					If FlagProc = True Then
						OrgCodi = OrgCodiNew
						PCodi = PCodiNew
						SCodi = SCodiNew
						ProcSolic = ProcSolicNew
						DtSolic = rsDiv("S" & m_oAIRData.RBAC & "_DTSTAT")
						TskIni = rsDiv("T" & m_oAIRData.RBAC & "_CODI")
						StOrigin = rsDiv( "TST_ORIGEM" ) ' Origem Solicitação 'A' ou 'E'
						StSolic = rsDiv("TST_CODI")
						DataPend   = ""
						DataAudit = ""
					End If

				End If

			Wend

		End If

		rsDiv.Close()

		ret = m_oAIRData.Write (dt, bExist)

		m_oAIRData.Close()

		GetValues = 1

	End Function

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
