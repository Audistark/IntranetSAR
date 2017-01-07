<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage = 1252 %> 
<!-- #include virtual = "/inet2/class/cCtrlErr.asp" -->
<!-- #include virtual = "/inet2/class/cLog.asp" -->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->

<!-- #include virtual = "/inet2/class/cAIRStatistics.asp" -->
<!-- #include virtual = "/inet2/class/cAIRStatsData.asp" -->

<%
Const ni = 3
Const nj = 3

Dim rbac : rbac = Array("121", "135", "145")
Dim gtar : gtar = Array("GTAR-DF", "GTAR-SP", "GTAR-RJ") 
Dim DtToday : DtToday = Date()
Dim wday : wday = Weekday(DtToday) ' sábado e domingo não
Dim tRefresh : tRefresh = 10 ' ten seconds
Dim tStamp : tStamp = 0
Dim nCount : nCount = 0
Dim i : i = -1
Dim j : j = -1
Dim ret : ret = 0
Dim strQuery : strQuery = Request.QueryString("nCount")

If wday = vbSunday Or wday = vbSaturday Then

	tRefresh = 14400 ' four hours

ElseIf strQuery <> "" Then

	tRefresh = 30 ' thirty seconds

	'-----------------------------------------------------------
	' Control Object
	Application.Lock
	tStamp = CDate( Application("AIRStatsTimestamp") )
	nCount = CInt( Application("AIRStatsCounter") )
	Dim hDiff : hDiff = DateDiff("h", tStamp, Now)
	If nCount < (ni * nj) Then
		If hDiff < 3 Then
			i = Fix( nCount / 3 )
			j = nCount mod 3
		Else
			Application("AIRStatsCounter") = 0
			nCount = 0
			i = 0
			j = 0
		End If
	ElseIf hDiff > 3 Then
		Application("AIRStatsCounter") = 0
		nCount = 0
		i = 0
		j = 0
	Else
		tRefresh = 3600 ' an hour
	End If 
	Application.UnLock
	'-------------------------------------------------------

	If i >= 0 And j >= 0 Then

		Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
		Dim querySQL, rsDiv, SolCodi
		Dim bSuccess : bSuccess = True

		querySQL =	"SELECT * FROM A" & rbac(i) & "_TabSolic ORDER By TSOL_CODI"
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)

		Do While Not rsDiv.Eof

			SolCodi = rsDiv( "TSOL_CODI" )

			Dim oAIRStats : Set oAIRStats = (new cAIRStatistics)( Array(rbac(i), gtar(j), SolCodi) )
			ret = oAIRStats.GetValues(DtToday)
			If ret < 0 Then
				bSuccess = False
			End If

			' Next
			rsDiv.MoveNext

		Loop

		'-----------------------------------------------------------
		' Control Object
		If bSuccess = True Then
			Application.Lock
			If nCount = Application("AIRStatsCounter") Then
				Application("AIRStatsTimestamp") = Now()
				Application("AIRStatsCounter") = nCount + 1
			End If
			Application.UnLock
		Else
			tRefresh = 600 ' 10 minutes
		End If
		'-------------------------------------------------------

		oDbFDH.Close()

	End If

End If

%>

<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="REFRESH" content="<%=tRefresh %>;URL=http://sar/inet2/wrAIRStatistics.asp?nCount=<%=nCount %>">
</head>
<body>
<pre>
i = <%=i %>, j = <%=j %>, nCount = <%=nCount %>
tStamp = <%=FormatDateTime(tStamp) %>
tRefresh = <%=tRefresh %>, ret = <%=ret %>
</pre>
</body>
</html>

