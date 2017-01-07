<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% Option Explicit %>
<% Response.CodePage = 1252 %>
<% Response.Buffer  = False %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/lib/libFuncDiv.asp" -->
<%

Dim Sec : Sec	= Request.QueryString( "Sec" )
Select Case Sec
	Case "145"
	Case "121"
	Case "135"
	Case Else
		Response.Status = "400 Bad Request"
		Response.Write "Argumentos inválidos."
		Response.End
End Select

' web service ?Srv=web.service
Dim WebSrv : WebSrv = Request.QueryString( "Srv" )
If WebSrv <> "web.service" Then
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If

Dim Row, Item, Value

Dim reset : reset = False

' command ?Command=fetch (/search)
Dim Command : Command = Request.QueryString( "Command" )
If Command = "fetch" Then
	Row = CLng("0" & Request.QueryString( "Row" ))
	If Row <= 0 Then
		Response.Status = "400 Bad Request"
		Response.Write "Argumentos inválidos."
		Response.End
	End If
ElseIf Command = "save" Then
	Row = CLng("0" & Request.QueryString( "Row" ))
	If Row <= 0 Then
		Response.Status = "400 Bad Request"
		Response.Write "Argumentos inválidos."
		Response.End
	End If
	Item = Request.QueryString( "Item" )
	Value = Request.QueryString( "Value" )
ElseIf Command = "range" Then
	Row = 1
ElseIf Command = "reset" Then
	reset = True
	Row = 1
Else
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If

'-----------------------------------------------------
'' Key
'Dim SecKey : SecKey = Request.QueryString("SEC_KEY")
'If Len(SecKey) > 5 Then
'	Response.Status = "400 Bad Request"
'	Response.End
'End If
'If Abs( CLng(SecKey) - Timer() ) > 60 Then
'	Response.Status = "400 Bad Request"
'	Response.End
'End If
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	Response.Status = "403 Forbidden"
	Response.End
End If

' Database access
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Sub PrintErr( err, descr )
%>
		<tr>
			<td>RET</td>
			<td><%=err %></td>
			<td><%=descr %></td>
		</tr>
	</table>
</body>
</html>
<%
	Response.End
End Sub

%>
<!DOCTYPE html>
<html>
<head>
  <title>AvGeralProcsWebService</title>
</head>
<body>
	<table border="1">
<%

Dim Gtar, Base, CHE, PCodi, PCodiDt, SolicSeq, TSolCodi, SolicDescr, SolicStatus, Analista
Dim TaskSeq, TaskDescr, TaskStatus, TaskStatusDt, TaskDt, TaskPendDt, TaskObs, Calc

Dim querySQL, rsDiv, res
Dim bCalculate : bCalculate = False

If Row = 1 Then
	If reset = True Then
		bCalculate = True
	Else
		' Verifica se tem que calcular mesmo
		querySQL = "SELECT TSTAMP FROM TMP_" & Sec & "PROCS WHERE ITEM = " & Row
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			Call PrintErr(-3, "Internal Error")
		End If
		If Not rsDiv.Eof Then
			' timestamp
			Dim tStamp : tStamp = rsDiv( "TSTAMP" )
			Dim minLast : minLast = DateDiff("n", tStamp, Now())
			If minLast < 120 Then ' não recalcula automaticamente!!!
				bCalculate = False
			Else
				bCalculate = True
			End If
		Else
			bCalculate = True
		End If
	End If
End If

If bCalculate Then ' recalculate

	querySQL = "DELETE * FROM TMP_" & Sec & "PROCS"
	ret = oDbFDH.Execute( querySQL )

	Dim OneYearAgo : OneYearAgo = DateAdd("m",-13,Date())

	querySQL = "  SELECT O.ORG_NABREV, B.CHE_CODI, P.P" & Sec & "_CODI, P.P" & Sec & "_DATA, " & _
			   "         P.P" & Sec & "_DOCMAN, S.S" & Sec & "_CODI, S.S" & Sec & "_DTSTAT, " & _
			   "         TS.TSOL_DESCR, TSt.TST_DESCR AS StatusSolic, TS.TSOL_CODI, Pes.PES_NGUERRA, " & _
			   "         Ger.SDIV_SIGLA, T.T" & Sec & "_CODI, TT.TSK_DESCR, T.T" & Sec & "_DATA, " & _
			   "         TSt_1.TST_DESCR AS StatusTarefa, T.T" & Sec & "_DTSTAT, " & _
			   "         T.T" & Sec & "_DTPEND, T.T" & Sec & "_OBS " & _
			   "    FROM ( ( ( ( ( ( ( ( ( Organizacao AS O INNER JOIN A" & Sec & "_Bases AS B " & _
			   "         ON O.ORG_CODI = B.ORG_CODI ) INNER JOIN A" & Sec & "_Processos AS P " & _
			   "         ON B.B" & Sec & "_CODI = P.B" & Sec & "_CODI ) " & _
			   "         INNER JOIN A" & Sec & "_Solicitacoes AS S " & _
			   "         ON P.P" & Sec & "_CODI = S.P" & Sec & "_CODI ) " & _
			   "         INNER JOIN A" & Sec & "_TabSolic AS TS ON S.TSOL_CODI = TS.TSOL_CODI ) " & _
			   "         LEFT JOIN A" & Sec & "_Tarefas AS T " & _
			   "         ON S.P" & Sec & "_S" & Sec & " = T.P" & Sec & "_S" & Sec & " ) " & _
			   "         LEFT JOIN A" & Sec & "_TabTarefa AS TT ON T.TSK_CODI = TT.TSK_CODI ) " & _
			   "         INNER JOIN Pessoal AS Pes ON S.PES_CODI = Pes.PES_CODI ) " & _
			   "         INNER JOIN Tab_Subdivisao AS Ger ON Pes.SDIV_CODI = Ger.SDIV_CODI ) " & _
			   "         INNER JOIN A" & Sec & "_TabStatus AS TSt ON S.TST_CODI = TSt.TST_CODI ) " & _
			   "         LEFT JOIN A" & Sec & "_TabStatus AS TSt_1 " & _
			   "         ON T.TST_CODI = TSt_1.TST_CODI " & _
			   " WHERE T.T" & Sec & "_CODI <> '' And TT.TSK_DESCR = 'Auditoria Técnica' And " & _
			   "		( TSt.TST_PROSSEGUE = 'S' OR S.S" & Sec & "_DTSTAT > #" & Month(OneYearAgo) & "/" & Day(OneYearAgo) & "/" & Year(OneYearAgo) & "# )" & _
	           "  ORDER BY Ger.SDIV_SIGLA, O.ORG_NABREV, P.P" & Sec & "_CODI, S.S" & Sec & "_CODI, T.T" & Sec & "_CODI"

	Dim i : i = 1

	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If
	If rsDiv.Eof Then
		Call PrintErr(-2, "Table Empty")
	End If

	Do

		Gtar = rsDiv( "SDIV_SIGLA" )
		Base = rsDiv( "ORG_NABREV" )
        CHE = rsDiv( "CHE_CODI" )
		PCodi = rsDiv( "P" & Sec & "_CODI" )
		If rsDiv( "P" & Sec & "_DATA" ) <> "" Then
			PCodiDt = Day( rsDiv( "P" & Sec & "_DATA" ) ) & "/" & _
					   Month( rsDiv( "P" & Sec & "_DATA" ) ) & "/" & _
						Year( rsDiv( "P" & Sec & "_DATA" ) )
		Else
			PCodiDt = ""
		End If
		SolicSeq = rsDiv( "S" & Sec & "_CODI" )
		TSolCodi = rsDiv( "TSOL_CODI" )
		SolicDescr = rsDiv( "TSOL_DESCR" )
		SolicStatus = rsDiv( "StatusSolic" )
		Analista = rsDiv( "PES_NGUERRA" )
		TaskSeq = rsDiv( "T" & Sec & "_CODI" )
		TaskDescr = rsDiv( "TSK_DESCR" )
		TaskStatus = rsDiv( "StatusTarefa" )
		If rsDiv( "T" & Sec & "_DATA" ) <> "" then
			TaskDt = Day( rsDiv( "T" & Sec & "_DATA" ) ) & "/" & _
					  Month( rsDiv( "T" & Sec & "_DATA" ) ) & "/" & _
					   Year( rsDiv( "T" & Sec & "_DATA" ) )
		Else
			TaskDt = ""
		End If
		If rsDiv( "T" & Sec & "_DTSTAT" ) <> "" Then
			TaskStatusDt = Day( rsDiv( "T" & Sec & "_DTSTAT" ) ) & "/" & _
							Month( rsDiv( "T" & Sec & "_DTSTAT" ) ) & "/" & _
							 Year( rsDiv( "T" & Sec & "_DTSTAT" ) )
		Else
			TaskStatusDt = ""
		End If
		If rsDiv( "T" & Sec & "_DTPEND" ) <> "" Then
			TaskPendDt = Day( rsDiv( "T" & Sec & "_DTPEND" ) ) & "/" & _
						  Month( rsDiv( "T" & Sec & "_DTPEND" ) ) & "/" & _
						   Year( rsDiv( "T" & Sec & "_DTPEND" ) )
		Else
			TaskPendDt = ""
		End If
		TaskObs = rsDiv( "T" & Sec & "_OBS" )

		Calc = ""

		On Error Resume Next
		querySQL =  "INSERT INTO TMP_" & Sec & "PROCS (ITEM, GTAR, NOMEABREV, CHE, PROCESSO, PROCESSODATA, " & _
					"SOLICSEQ, TSOLCODI, SOLICDESCR, SOLICSTATUS, ANALISTA, TASKSEQ, TASKDESCR, TASKSTATUS, " & _
					"TASKSTATUSDATA, TASKDATA, TASKSPENDDATA, TASKOBSERV, CALC, TSTAMP) " & _
					"VALUES (" & i & ", '" & Gtar & "', '" & Base & "', '" & CHE & "', '" & PCodi & "', '" & PCodiDt & "', " & _
					SolicSeq & ", '" & TSolCodi & "', '" & SolicDescr & "', '" & SolicStatus & "', '" & Analista & "', " & _
					TaskSeq & ", '" & TaskDescr & "', '" & TaskStatus & "', '" & TaskStatusDt & "', '" & _
					TaskDt & "', '" & TaskPendDt & "', '" & Replace(TaskObs,"'","''") & "', '" & Calc & "', '" & Now() & "');"
		ret = oDbFDH.Execute( querySQL )
		If Err.Number = 0 Then
			i = i + 1
		Else
			querySQL = "DELETE * FROM TMP_" & Sec & "PROCS"
			ret = oDbFDH.Execute( querySQL )
			Call PrintErr(-Err.Number, Err.Description)
		End If
		On Error GoTo 0

		rsDiv.MoveNext

	Loop While Not rsDiv.Eof 

End If
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'------------------------------------------------------
'
'	Fetch procs
'
If Command = "fetch" Or Command = "reset" Then

	querySQL = "SELECT * FROM TMP_" & Sec & "PROCS WHERE ITEM >= " & Row & " AND CALC <> '1' ORDER BY ITEM"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If
	If rsDiv.Eof Then
		Call PrintErr(0, "Not Found")
	End If

'------------------------------------------------------
'
'	Set value
'
ElseIf Command = "save" Then

	Dim strProcs(19) : strProcs(1) = "ITEM" : strProcs(2) = "GTAR" : strProcs(3) = "NOMEABREV" : strProcs(4) = "CHE" : _
					   strProcs(5) = "PROCESSO" : strProcs(6) = "PROCESSODATA" : strProcs(7) = "SOLICSEQ" : _
					   strProcs(8) = "TSOLCODI" : strProcs(9) = "SOLICDESCR" : strProcs(10) = "SOLICSTATUS" : _
					   strProcs(11) = "ANALISTA" : strProcs(12) = "TASKSEQ" : strProcs(13) = "TASKDESCR" : _
					   strProcs(14) = "TASKSTATUS" : strProcs(15) = "TASKSTATUSDATA" : strProcs(16) = "TASKDATA" : _
					   strProcs(17) = "TASKSPENDDATA" : strProcs(18) = "TASKOBSERV" : strProcs(19) = "CALC"
 
	' Le em busca do cara
	querySQL = "UPDATE TMP_" & Sec & "PROCS SET " & strProcs(Item) & "='" & Value & "' WHERE ITEM=" & Row
	oDbFDH.Execute(querySQL)

	Call PrintErr(1, "Operation successfully performed.")

'------------------------------------------------------
'
'	Procs UsedRangeRowsCount
'
ElseIf Command = "range" Then

	querySQL = "SELECT COUNT(*) AS Count FROM TMP_" & Sec & "PROCS WHERE CALC <> '1'"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If
	If rsDiv.Eof Then
		Call PrintErr(0, "Not Found")
	End If

	Call PrintErr(rsDiv("Count"), "Operation successfully performed.")

Else

	Call PrintErr(-1, "Argumentos inválidos.")

End If

%>
		<tr>
			<td>RET</td>
			<td>1</td>
		</tr>
		<tr>
			<td>ITEM</td>
			<td><%=rsDiv("ITEM") %></td>
		</tr>
		<tr>
			<td>GTAR</td>
			<td><%=rsDiv("GTAR") %></td>
		</tr>
		<tr>
			<td>NOMEABREV</td>
			<td><%=rsDiv("NOMEABREV") %></td>
		</tr>
		<tr>
			<td>CHE</td>
			<td><%=rsDiv("CHE") %></td>
		</tr>
		<tr>
			<td>PROCESSO</td>
			<td><%=rsDiv("PROCESSO") %></td>
		</tr>
		<tr>
			<td>PROCESSODATA</td>
			<td><%=rsDiv("PROCESSODATA") %></td>
		</tr>
		<tr>
			<td>SOLICSEQ</td>
			<td><%=rsDiv("SOLICSEQ") %></td>
		</tr>
		<tr>
			<td>TSOLCODI</td>
			<td><%=rsDiv("TSOLCODI") %></td>
		</tr>
		<tr>
			<td>SOLICDESCR</td>
			<td><%=rsDiv("SOLICDESCR") %></td>
		</tr>
		<tr>
			<td>SOLICSTATUS</td>
			<td><%=rsDiv("SOLICSTATUS") %></td>
		</tr>
		<tr>
			<td>ANALISTA</td>
			<td><%=rsDiv("ANALISTA") %></td>
		</tr>
		<tr>
			<td>TASKSEQ</td>
			<td><%=rsDiv("TASKSEQ") %></td>
		</tr>
		<tr>
			<td>TASKDESCR</td>
			<td><%=rsDiv("TASKDESCR") %></td>
		</tr>
		<tr>
			<td>TASKSTATUS</td>
			<td><%=rsDiv("TASKSTATUS") %></td>
		</tr>
		<tr>
			<td>TASKSTATUSDATA</td>
			<td><%=rsDiv("TASKSTATUSDATA") %></td>
		</tr>
		<tr>
			<td>TASKDATA</td>
			<td><%=rsDiv("TASKDATA") %></td>
		</tr>
		<tr>
			<td>TASKSPENDDATA</td>
			<td><%=rsDiv("TASKSPENDDATA") %></td>
		</tr>
		<tr>
			<td>TASKOBSERV</td>
			<td><%=rsDiv("TASKOBSERV") %></td>
		</tr>
		<tr>
			<td>CALC</td>
			<td><%=rsDiv("CALC") %></td>
		</tr>
		<tr>
			<td>TIMESTAMP</td>
			<td><%=rsDiv( "TSTAMP" ) %></td>
		</tr>
	</table>
</body>
</html>
<%

rsDiv.Close
oDbFDH.Close

 %>
