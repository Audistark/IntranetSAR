<%

'----------------------------------------------------------------
'
'	Class cStCalc
'
'	Date: 03/05/2015
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------


'----------------------------------------------------------------
'
'	Class cStCalc
'
Class cStCalc

	' Declarations
	Private m_oCtrlErr			' Error Object
	Private m_oAIRData			' Data Object

	' Class Initialization
	Private Sub Class_Initialize()
		Set m_oCtrlErr	= new cCtrlErr
		Set m_oAIRData	= new cStData
	End Sub
	Public Default Function construct( parameters )
		' "145", "GTAR-DF", "S", "001", 15
		Dim ret : ret = UBound(parameters)
		If ret < 4 Then
			m_oCtrlErr.Error = "Invalid Arguments."
		Else
			m_oAIRData.RBAC  = parameters(0)
			m_oAIRData.GTAR	 = parameters(1)
			m_oAIRData.sType = parameters(2)
			m_oAIRData.Codi  = parameters(3)
			m_oAIRData.Goal  = parameters(4)
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

	' Goal
	Public Property Get Goal()
		Goal = m_oAIRData.Goal
	End Property

	' Open Connection
	Public Function GetValues()

		If m_oAIRData.Open() < 0 then
			m_oCtrlErr.Import( m_oAIRData.getObjErr() )
			GetValues = m_oCtrlErr.ErrorNumber
			Exit Function
		End If

		Dim ret : ret = -1
		Dim bExist : bExist = False
		Dim docman_ant : docman_ant = ""

		Dim tStamp : tStamp = Now()
		Dim dtToday : dtToday = Date()

		'--------------------------------------------------------------
		' Verifica se j� foi criado algum registro para a data desejada
		Dim try : try = -1
		Do
			ret = m_oAIRData.Read(dtToday)
			If ret < 0 Then
				m_oCtrlErr.Import( m_oAIRData.getObjErr() )
				GetValues = m_oCtrlErr.ErrorNumber
				Exit Function
			End If
			If ret > 0 Then

				' Yes it exists
				bExist = True

				' Se n�o � hoje n�o precisa recalcular
				If dtToday <> Date() Then
					m_oAIRData.Close()
					GetValues = 0
					Exit Function
				End If

				' Previne overrun (3h)
				ret = DateDiff("h", m_oAIRData.tStamp, Now())
				If ret < 3 Then
					m_oAIRData.Close()
					GetValues = 0
					Exit Function
				End If

				Exit Do ' for�a saida (jump)

			Else ' not found

				' Se n�o existe por�m � s�bado ou domingo...
				Dim wday : wday = Weekday(dtToday) ' s�bado e domingo n�o
				If wday = vbSunday Or wday = vbSaturday Then
					If try < 0 Then
						try = 5 ' tenta voltar at� 5 dias
					ElseIf try = 0 Then
						GetValues = -1 ' n�o encontrou registro
						Exit Function
					Else
						try = try - 1
					End If
					dtToday = DateAdd("d", -1, dtToday)
				Else

					Exit Do ' for�a saida (jump)

				End If

			End If

		Loop

		' Calcula para a data de hoje

		m_oAIRData.CleanValues()

		Dim oDbFDH, rsDiv, querySQL
		Set oDbFDH = m_oAIRData.oDbFDH

		'
		' Busca valores
		'
		Dim sSolCodi: sSolCodi = "000"
		Dim sTskCodi: sTskCodi = "000"
		If m_oAIRData.sType = "S"  Then
			sSolCodi = m_oAIRData.Codi
		Else
			sTskCodi = m_oAIRData.Codi
		End If
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
					"    WHERE ( TS.TSOL_CODI='" & sSolCodi & "' OR TT.TSK_CODI='" & sTskCodi & "' ) AND " & _
					"          TSDiv.SDIV_SIGLA='" & m_oAIRData.GTAR & "' AND O.ORG_BR='S' AND " & _
					"          ( TSt.TST_PROSSEGUE='S' OR DateDiff( 'd', S.S" & m_oAIRData.RBAC & "_DTSTAT, Now) < 32 ) " & _
					"    ORDER BY P.P" & m_oAIRData.RBAC & "_CODI, S.S" & m_oAIRData.RBAC & "_CODI, T.T" & m_oAIRData.RBAC & "_CODI, D.D" & m_oAIRData.RBAC & "_DATA" 
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)

		If rsDiv Is Nothing Then
			m_oCtrlErr.Import( oDbFDH.getObjErr() )
			m_oAIRData.Close()
			GetValues = m_oCtrlErr.ErrorNumber
			Exit Function
		End If

		If Not rsDiv.Eof Then

			Dim FlagProc : FlagProc = True		' Processo
			Dim FlagEOReg : FlagEOReg = False	' 
			Dim PCodi : PCodi = rsDiv( "P" & m_oAIRData.RBAC & "_CODI" )
			Dim SCodi : SCodi = rsDiv( "S" & m_oAIRData.RBAC & "_CODI" )
			Dim StOrigin : StOrigin = rsDiv( "TST_ORIGEM" ) ' Origem Solicita��o 'A' ou 'E'
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
				If rsDiv("SOLIC_NCLOSED") = "S" Then ' N�o conclu�do
					bSolicProssegue = True
				End If

				Dim bStProssegue : bStProssegue = False
				If rsDiv("TASK_NCLOSED") = "S" Then ' N�o conclu�do
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

				' Verifica emiss�o de documentos pela ANAC nos �ltimos 30 dias
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

				' procura por data de pend�ncia mais pr�xima em tarefas n�o concluidas..
				' e atualiza a data da solicita��o
				If bStProssegue = True And bContaTempo = True Then ' N�o conclu�do e conta tempo

					' Se tem auditoria agendada no futuro.. 
					If DtTaskExec <> "" And bTaskDate = True And _
					   ( DataAudit = "" Or DtTaskExec > DataAudit ) Then
						If ( DtTaskExec + 10 ) >= Date() Then
							DataAudit = DtTaskExec
						End If
					End If

					' procura pela data de pend�ncia mais nova..
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
					If OrgCodi <> OrgCodiNew Or ProcSolic <> ProcSolicNew Then
						FlagEOReg = True
					End If
				End If

				' Finaliza C�lculos
				If FlagEOReg = True Then

					' Data Solic devido a Auditoria
					If StOrigin = "A" And _
						DataAudit <> "" And _
						DataAudit > DtSolic Then
						DtSolic = DataAudit
					End If

					' Verifica solicita��es fechadas nos �ltimos 30 dias
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

						' ANAC
						If StOrigin = "A" Then

							m_oAIRData.ANAC = m_oAIRData.ANAC + 1

							If Delay > m_oAIRData.ANACMaxDays Then
								m_oAIRData.ANACMaxDays = Delay
							End If
							
							If ( m_oAIRData.Goal > 0 And Delay > m_oAIRData.Goal ) Or _
								Delay > 60 Then ' 60d is the default goal
								m_oAIRData.ANACDelay = m_oAIRData.ANACDelay + 1
								Call m_oAIRData.TopTen( ProcSolic, Delay )
							End If

							If Delay > 0 Then
								m_oAIRData.ANACDays = m_oAIRData.ANACDays + Delay
							End If

						' Empresa
						Else

							m_oAIRData.Client = m_oAIRData.Client + 1
						
							If DataPend <> "" Then
								If Delay > 30 Then		' D� uma folga no prazo da pend�ncia
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
						StOrigin = rsDiv( "TST_ORIGEM" ) ' Origem Solicita��o 'A' ou 'E'
						StSolic = rsDiv("TST_CODI")
						DataPend   = ""
						DataAudit = ""
					End If

				End If

			Wend

		End If

		rsDiv.Close()

		ret = m_oAIRData.Write (dtToday, bExist)

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
