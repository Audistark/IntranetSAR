<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.CodePage = 1252 %> 
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->

<!-- #include virtual = "/inet2/wd/hSolicConsts.asp" -->

<!-- #include virtual = "/inet2/wd/cWDInclEO145.asp" -->
<!-- #include virtual = "/inet2/wd/cWDManuals145.asp" -->
<!-- #include virtual = "/inet2/wd/cWDCert145.asp" -->
<!-- #include virtual = "/inet2/wd/cWDOthers145.asp" -->
<!-- #include virtual = "/inet2/wd/cWDInclEO135.asp" -->
<!-- #include virtual = "/inet2/wd/cWDManuals135.asp" -->
<!-- #include virtual = "/inet2/wd/cWDMEL135.asp" -->
<!-- #include virtual = "/inet2/wd/cWDPM135.asp" -->
<!-- #include virtual = "/inet2/wd/cWDOthers135.asp" -->
<!-- #include virtual = "/inet2/wd/cWDRCA91.asp" -->

<!-- #include virtual = "/inet2/wd/cAIRStatistics.asp" -->
<!-- #include virtual = "/inet2/wd/cAIRStatsData.asp" -->
<!-- #include virtual = "/inet2/wd/cAIRStatsRules.asp" -->

<%
Dim gtar : gtar = Request.QueryString("gtar")
If gtar <> "GTAR-DF" And gtar <> "GTAR-SP" And gtar <> "GTAR-RJ" Then
	Response.Status = "400 Bad Request"
	Response.Write "Invalid Arguments."
	Response.End
End If

Dim rbac : rbac = Request.QueryString("rbac")
If rbac <> "91" And rbac <> "135" And rbac <> "145" Then
	Response.Status = "400 Bad Request"
	Response.Write "Invalid Arguments."
	Response.End
End If

' 145
Dim oWDInclEO145 : Set oWDInclEO145 = (new cWDInclEO145)(gtar)
Dim oWDManuals145 : Set oWDManuals145 = (new cWDManuals145)(gtar)
Dim oWDCert145 : Set oWDCert145 = (new cWDCert145)(gtar)
Dim oWDOthers145 : Set oWDOthers145 = (new cWDOthers145)(gtar)

' 135
Dim oWDInclEO135 : Set oWDInclEO135 = (new cWDInclEO135)(gtar)
Dim oWDMEL135 : Set oWDMEL135 = (new cWDMEL135)(gtar)
Dim oWDPM135 : Set oWDPM135 = (new cWDPM135)(gtar)
Dim oWDManuals135 : Set oWDManuals135 = (new cWDManuals135)(gtar)
Dim oWDOthers135 : Set oWDOthers135 = (new cWDOthers135)(gtar)

' 91
Dim ocWDRCA91 : Set ocWDRCA91 = (new cWDRCA91)(gtar)


%>
<!DOCTYPE html>
<html>
<head>

  <title>WatchDog/<%=gtar %>/RBAC<%=rbac %></title>

  <style>
	.linha
	{
		border-bottom: 1px solid black;
		margin: 2px 12px 7px 12px;
	}
	.linha2
	{
		border-bottom: 1px dashed #c9cacb;
		margin: 7px 12px 7px 12px;
		width: 95%;
	}
  </style>
  
  <!--Load the AJAX API-->
  <script type="text/javascript" src="https://www.google.com/jsapi"></script>

  <script type="text/javascript">
	function refreshEvery3Hours() {
		var now = new Date();
		if (now.getHours() % 3 == 0 && now.getMinutes() < 2 ) {
			window.location.reload(true);
		}
	}
	var myCheck = setInterval(function () { refreshEvery3Hours() }, 60000);
  </script>
	<%
	Select Case rbac

		Case "91"

			ocWDRCA91.JavaScripts()
			ocWDRCA91.PrintGauge()
			ocWDRCA91.PrintCharts()
			ocWDRCA91.PrintPieChart()

		Case "135"

			oWDInclEO135.JavaScripts()
			oWDInclEO135.PrintGauge()
			oWDInclEO135.PrintCharts()
			oWDInclEO135.PrintPieChart()

			oWDMEL135.JavaScripts()
			oWDMEL135.PrintGauge()
			oWDMEL135.PrintCharts()
			oWDMEL135.PrintPieChart()

			oWDPM135.JavaScripts()
			oWDPM135.PrintGauge()
			oWDPM135.PrintCharts()
			oWDPM135.PrintPieChart()

			oWDManuals135.JavaScripts()
			oWDManuals135.PrintGauge()
			oWDManuals135.PrintCharts()
			oWDManuals135.PrintPieChart()

			oWDOthers135.JavaScripts()
			oWDOthers135.PrintGauge()
			oWDOthers135.PrintCharts()
			oWDOthers135.PrintPieChart()

		Case "145"

			oWDInclEO145.JavaScripts()
			oWDInclEO145.PrintGauge()
			oWDInclEO145.PrintCharts()
			oWDInclEO145.PrintPieChart()

			oWDManuals145.JavaScripts()
			oWDManuals145.PrintGauge()
			oWDManuals145.PrintCharts()
			oWDManuals145.PrintPieChart()

			oWDCert145.JavaScripts()
			oWDCert145.PrintGauge()
			oWDCert145.PrintPieChart()

			oWDOthers145.JavaScripts()
			oWDOthers145.PrintGauge()
			oWDOthers145.PrintPieChart()

	End Select %>

</head>
<body>

  <div class="linha"></div>

  <table border="0" cellpadding="0" cellspacing="0" width="98%">
    <tr align="center" valign="middle">
        <td><img src=/inet2/img/watchdog.jpg border=0 width=82 height=82></td>
        <td align="center"><font size="8" face="Calibri"><strong>WatchDog - SAR/<%=gtar %>/RBAC <%=rbac %></strong></font></td>
		<td align="right" valign=bottom><font size="3" face="Calibri">Date: <%=Now() %><br>Version 1.0<br>Grupo de TI/SAR</font></td>
    </tr>
  </table>

  <div class="linha"></div>

	<%
	Select Case rbac

		Case "145" %>

		<%	oWDInclEO145.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDManuals145.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDCert145.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDOthers145.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%
		Case "135" %>

		<%	oWDInclEO135.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDMEL135.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDPM135.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDManuals135.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%	oWDOthers135.PrintHtml() %>

  <center><div class="linha2"></div></center>

		<%
		Case "91" %>

		<%	ocWDRCA91.PrintHtml() %>

  <center><div class="linha2"></div></center>

	<%
	End Select %>

</body>
</html>
