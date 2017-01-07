<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage = 1252 %> 
<% session.LCID = 1046 'BRASIL %>

<!-- #include virtual = "/inet2/class/cCtrlErr.asp" -->
<!-- #include virtual = "/inet2/class/cLog.asp" -->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->

<!-- #include virtual = "/inet2/stats/cStatistics.asp" -->
<!-- #include virtual = "/inet2/stats/cStCalc.asp" -->
<!-- #include virtual = "/inet2/stats/cStData.asp" -->

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

If wday = vbSaturday Or wday = vbSunday Then

	tRefresh = 14400 ' four hours

ElseIf strQuery <> "" Then

	tRefresh = 30 ' thirty seconds

	'-----------------------------------------------------------
	' Control Object
	Application.Lock
	tStamp = CDate( Application("AIRStatsTimestamp") )
	nCount = CInt( Application("AIRStatsCounter") )

''' Force calc ''''''''''
' tStamp = Now()
' nCount = 0
' Application("AIRStatsTimestamp") = DateAdd("y", 2, Now())
'''''''''''''''''''''''''

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

'''' 145 '''''
'i = 2
'j = 2
''''''''''''''

		Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
		Dim querySQL, rsDiv
		Dim bSuccess : bSuccess = True

		querySQL =	"SELECT * FROM A" & rbac(i) & "_TabStatsGroup ORDER By TStatsGr_Id"
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)

		Do While Not rsDiv.Eof

			Dim iGroup : iGroup = rsDiv( "TStatsGr_Id" )

			Dim oStats : Set oStats = (new cStatistics)( Array(rbac(i), gtar(j), iGroup) )
			If oStats.Ret() < 0 Then
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

