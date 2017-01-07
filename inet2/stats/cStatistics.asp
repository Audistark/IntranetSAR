<%

'----------------------------------------------------------------
'
'	Class cStatistics
'
'	Date: 02/05/2015
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------
'
Const MaxItems = 10

'----------------------------------------------------------------
'	Results:
'
'	1- Ótimo			GREEN
'	2- Bom				NAVY
'	3- Alerta			YELLOW
'	4- Ruim				BROW
'	5- Crítico			RED BLINK	  Gauge
Const RES_NONE			= 0			 '   0
Const RES_OTIMO			= 1			 '  20
Const RES_BOM			= 2			 '  40
Const RES_ALERTA		= 3			 '  60
Const RES_RUIM			= 4			 '  80
Const RES_CRITICO		= 5			 ' 100

Const MaxDays			= 90

'----------------------------------------------------------------
'
'	Class cStatistics
'
'
'	Monta a visualização de um grupo estatístico cadastrado
'
'	Set oStats = (new cStatistics)( Array("145", "GTAR-DF", 1) )
'
'

Class cStatistics

	'Declarations
	Private m_oAIRCalc(30)		' AIRStatistics Calculation
	Private m_nAIRCalc			' number of objects

	Private m_oAIRData			' Data Object
	
	Private m_oDataHistory(90)	' Data Object with a maximum of three months
	Private m_nDataHistory

	Private m_ret				' return value

	Private	i, j, k

	' Variables
	Private m_strGTAR			' GTAR (GTAR-DF, GTAR-SP, GTAR-RJ ou SAR)
	Private m_nGroup			' Group
	Private m_RBAC				' RBAC (121, 135 ou 145)
	Private m_bSend				' True if send statistical e-mail

	Private m_bodyMail			' Corpo do email

	Private m_sTitle			' Nome do Grupo
	Private m_iGoal				' Goal

	Private m_sObject			' Object Name

	' Solicitações
	Private m_TSolCodi(10)
	Private m_TSolDescr(10)
	Private m_nTSol

	' Tarefas
	Private m_TTskCodi(10)
	Private m_TTskDescr(10)
	Private m_nTTsk


	'Class Initialization
	Private Sub Class_Initialize()
		Set m_oAIRData	= new cStData
		For i=0 To MaxDays-1
			Set m_oDataHistory(i) = new cStData
		Next
		m_nAIRCalc = 0
		m_bodyMail = ""
		m_ret = 1
	End Sub
	Public Default Function construct( parameters )
		' "145", "GTAR-DF", 1, True
		Dim ret : ret = UBound(parameters)
		If ret < 2 Then
			m_oCtrlErr.Error = "Invalid Arguments."
			m_ret = -1
			Exit Function
		Else
			m_RBAC = parameters(0)
			m_strGTAR = UCase(parameters(1))
			m_nGroup = parameters(2)
			m_bSend = False
			If ret > 2 Then m_bSend = parameters(3)
		End If
        Set construct = Me

		' sObject
		m_sObject = m_nGroup
		Dim posGtar : posGtar = InStr( m_strGTAR, "-" )
		If posGtar > 0 Then ' GTAR-XX
			m_sObject = m_nGroup & Mid(m_strGTAR,posGtar+1,Len(m_strGTAR)-posGtar) & m_RBAC
		Else
			m_sObject = m_nGroup & m_strGTAR & m_RBAC
		End If

		m_nTSol = 0
		m_nTTsk = 0

		Dim querySQL, rsDiv
		Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
		If oDbFDH.ErrorNumber < 0 then
			m_ret = -1
			oDbFDH.Print()
		End If

		' Read A999_TabStatsGroup 
		querySQL = "SELECT * FROM A" & m_RBAC & "_TabStatsGroup Where TStatsGr_Id = " & m_nGroup
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			m_ret = -1
			oDbFDH.Print()
		End If
		If rsDiv.Eof Then
			m_ret = -1
			Exit Function
		End If
			
		m_sTitle = rsDiv( "TStatsGr_Name" )
		
		m_iGoal = rsDiv( "TStatsGr_Goal" )
		rsDiv.Close()

		Dim dt : dt = Date()
		
		m_oAIRData.dtDate = dt
		m_oAIRData.GTAR = m_strGTAR
		m_oAIRData.RBAC = m_RBAC
		m_oAIRData.Goal = m_iGoal

		' Get TSOL_CODI 
		querySQL = "SELECT * FROM A" & m_RBAC & "_TabSolic Where TSOL_GROUP = " & m_nGroup & " ORDER BY TSOL_CODI"
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			m_ret = -1
			oDbFDH.Print()
		End If
		On Error Resume Next
		If Err.Number = 0 Then
			Do While not rsDiv.eof
				m_TSolCodi(m_nTSol) = rsDiv("TSOL_CODI")
				m_TSolDescr(m_nTSol) = rsDiv("TSOL_DESCR")
				If m_nTSol < MaxItems Then
					m_nTSol = m_nTSol + 1
				End If
				rsDiv.MoveNext
			Loop
		End If
		On Error GoTo 0
		rsDiv.Close()

		' Get TSK_CODI 
		querySQL = "SELECT * FROM A" & m_RBAC & "_TabTarefa Where TSK_GROUP = " & m_nGroup & " ORDER BY TSK_CODI"
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			m_ret = -1
			oDbFDH.Print()
		End If
		On Error Resume Next
		If Err.Number = 0 Then
			Do While not rsDiv.eof
				m_TTskCodi(m_nTTsk) = rsDiv("TSK_CODI")
				m_TTskDescr(m_nTTsk) = rsDiv("TSK_DESCR")
				If m_nTTsk < MaxItems Then
					m_nTTsk = m_nTTsk + 1
				End If
				rsDiv.MoveNext
			Loop
		End If
		On Error GoTo 0
		rsDiv.Close()

		Dim tGtar : tGtar = Array("GTAR-DF", "GTAR-SP", "GTAR-RJ")

		For i=0 To m_nTSol-1
			If posGtar <= 0 Then ' SAR
				For k=0 To 2
					Set m_oAIRCalc(m_nAIRCalc) = (new cStCalc)( Array(m_RBAC, tGtar(k), "S", m_TSolCodi(i), m_iGoal) )
					If m_oAIRCalc(m_nAIRCalc).GetValues() < 0 Then
						m_ret = -1
						Exit Function
					End If
					m_oAIRData.Add(m_oAIRCalc(m_nAIRCalc).oData)
					m_nAIRCalc = m_nAIRCalc + 1
				Next
			Else
				Set m_oAIRCalc(m_nAIRCalc) = (new cStCalc)( Array(m_RBAC, m_strGTAR, "S", m_TSolCodi(i), m_iGoal) )
				If m_oAIRCalc(m_nAIRCalc).GetValues() < 0 Then
					m_ret = -1
					Exit Function
				End If
				m_oAIRData.Add(m_oAIRCalc(m_nAIRCalc).oData)
				m_nAIRCalc = m_nAIRCalc + 1
			End If
		Next
		For j=0 To m_nTTsk-1
			If posGtar <= 0 Then ' SAR
				For k=0 To 2
					Set m_oAIRCalc(m_nAIRCalc) = (new cStCalc)( Array(m_RBAC, tGtar(k), "T", m_TTskCodi(i), m_iGoal) )
					If m_oAIRCalc(m_nAIRCalc).GetValues() < 0 Then
						m_ret = -1
						Exit Function
					End If
					m_oAIRData.Add(m_oAIRCalc(m_nAIRCalc).oData)
					m_nAIRCalc = m_nAIRCalc + 1
				Next
			Else
				Set m_oAIRCalc(m_nAIRCalc) = (new cStCalc)( Array(m_RBAC, m_strGTAR, "T", m_TTskCodi(j), m_iGoal) )
				If m_oAIRCalc(m_nAIRCalc).GetValues() < 0 Then
					m_ret = -1
					Exit Function
				End If
				m_oAIRData.Add(m_oAIRCalc(m_nAIRCalc).oData)
				m_nAIRCalc = m_nAIRCalc + 1
			End If
		Next

		' empty
		If m_nTSol + m_nTTsk = 0 Then
			m_ret = 0
		End If

        Set construct = Me

    End Function

	'Terminate Class
	Private Sub Class_Terminate()
	End Sub

	Public Function Ret()
		Ret = m_ret
	End Function

	Public Function GetMail()
		GetMail = m_bodyMail
	End Function

	Public Function GetObjId()
		GetObjId = m_sObject
	End Function

	Public Function PreparedMail()

		PreparedMail = True

	End Function

	Public Function IsPriority()
		If m_iGoal > 0 Then
			IsPriority = True
		Else
			IsPriority = False
		End If
	End Function

	Public Function Alarm()
		Alarm = False
		If m_iGoal > 0 And m_oAIRData.ANACMaxDays > m_iGoal Then
			Alarm = True
		End If
	End Function
	
	Public Sub JavaScripts() %>
	<!-- #include virtual = "/inet2/stats/hStIncl.asp" -->
	<%
	End Sub


	Public Sub PieChart() %>
  <script type="text/javascript">
	google.load("visualization", "1", {packages:["corechart"]});
	google.setOnLoadCallback(drawPieChart<%=m_sObject %>);
	function drawPieChart<%=m_sObject %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value'],
			['ANAC', <%=m_oAIRData.ANAC %>],
			['CLIENTE', <%=m_oAIRData.Client %>],
			['CLOSED 30d', <%=m_oAIRData.Closed30d %>]
			]);
		var options = {
			backgroundColor: {fill:'white'},
			height: 120, width: 240,
			pieSliceText: 'value',
			enableInteractivity: true,
			chartArea: {left:0,top:6,width:'100%',height:'90%'}
		};
		var chart_div = document.getElementById('piechart_<%=m_sObject %>');
		var chart = new google.visualization.PieChart(chart_div);
		chart.draw(data, options);
	}
  </script>

	<%
	End Sub


	Public Sub Gauge()

		'--------------------------------------------------------------------------
		' resDelayANAC	- Indica atraso da ANAC na analise dos processos
		Dim resDelayANAC
		Dim goal : goal = m_oAIRData.Goal
		If goal <= 0 Then goal = 60
		If m_oAIRData.ANAC = 0 Then
			resDelayANAC = RES_NONE
		' CRITICAL
		ElseIf m_oAIRData.ANACDelay > 1 Or _
				m_oAIRData.ANACMaxDays > 2*goal Then
			resDelayANAC = RES_CRITICO
		' RES_RUIM
		ElseIf m_oAIRData.ANACDelay > 0 Then
			resDelayANAC = RES_RUIM
		' RES_ALERTA
		ElseIf m_oAIRData.ANACMaxDays > 0.8*goal Then
			resDelayANAC = RES_ALERTA
		' RES_BOM
		ElseIf m_oAIRData.ANACMaxDays > 0.4*goal Then
			resDelayANAC = RES_BOM
		' RES_OTIMO
		Else
			resDelayANAC = RES_OTIMO
		End If
		'--------------------------------------------------------------------------
		' resDelayClient	- Indica atraso da Empresa
		Dim resDelayClient
		If m_oAIRData.Client = 0 Then
			resDelayClient = RES_NONE
		' CRITICAL
		ElseIf m_oAIRData.ClientDelay > 5 Or _
				m_oAIRData.ClientMaxDays > 180 Then
			resDelayClient = RES_CRITICO
		' RES_RUIM
		ElseIf m_oAIRData.ClientDelay > 3 Or _
				m_oAIRData.ClientMaxDays > 120 Then
			resDelayClient = RES_RUIM
		' RES_ALERTA
		ElseIf m_oAIRData.ClientDelay > 2 Or _
				m_oAIRData.ClientMaxDays > 90 Then
			resDelayClient = RES_ALERTA
		' RES_BOM
		ElseIf m_oAIRData.ClientDelay > 1 Or _
				m_oAIRData.ClientMaxDays > 60 Then
			resDelayClient = RES_BOM
		' RES_OTIMO
		Else
			resDelayClient = RES_OTIMO
		End If
	 %>
  <script type="text/javascript">
  	google.load('visualization', '1', { packages: ['gauge'] });
	google.setOnLoadCallback(drawGauges<%=m_sObject %>);
	function drawGauges<%=m_sObject %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value'],
			['ANAC', 0],
			['CLIENTE', 0]
			]);
		var options = {
			width: 270, height: 120,
			redFrom: 85, redTo: 100,
			yellowFrom: 60, yellowTo: 85,
			minorTicks: 5
		};
		var gauge_div = document.getElementById('gauge_<%=m_sObject %>');
		var gauge = new google.visualization.Gauge(gauge_div);
		setInterval(function() {
			var val=Math.floor((<%=resDelayANAC %>-0.50)*20 + 15*Math.random())
			if( val > 100 ) val = 100
			if( val <= 25 ) val = Math.floor(val*0.75)
			if( val < 0 ) val = 0
			data.setValue(0, 1, val);
			gauge.draw(data, options);
		}, 5000);
		setInterval(function() {
			var val=Math.floor((<%=resDelayClient %>-0.50)*20 + 15*Math.random())
			if( val > 100 ) val = 100
			if( val <= 25 ) val = Math.floor(val*0.75)
			if( val < 0 ) val = 0
			data.setValue(1, 1, val);
			gauge.draw(data, options);
		}, 5500);
		gauge.draw(data, options);
	}
  </script>
	<%
	End Sub


	Public Sub Charts()

		Dim first
		Dim max : max = 10

		m_nDataHistory = calcHistory()

	%>
	<script type="text/javascript"><%
	If m_bSend = True And m_iGoal > 0 Then %>
	var chk1<%=m_sObject %> = false;
	var chk2<%=m_sObject %> = false;<%
	End If %>
	google.load('visualization', '1', { packages: ['corechart'] });
	google.setOnLoadCallback(drawChart1<%=m_sObject %>);
	function drawChart1<%=m_sObject %>() {
		// Create and populate the data table.
		var data = new google.visualization.DataTable();
		data.addColumn('date', 'Date');
		data.addColumn('number', 'ANAC');
		data.addColumn('number', 'CLIENTE');
		data.addColumn('number', 'ANAC DELAY');
		data.addColumn('number', 'CLIENT DELAY');
		data.addColumn('number', '30d DOCS');
		data.addColumn('number', '30d CLOSED');
		data.addRows([
		<%
			first = True
			Dim last : last = 30
			If last > m_nDataHistory Then last = m_nDataHistory
			For i=last-1 To 0 Step -1
				If Not first Then Response.Write("," & vbCrLf)
				first = False
				Response.Write( "	[new Date(" & Year(m_oDataHistory(i).dtDate) & ", " & Month(m_oDataHistory(i).dtDate) - 1 & ", " & Day(m_oDataHistory(i).dtDate) & "), " & _
								m_oDataHistory(i).ANAC & ", " & _
								m_oDataHistory(i).Client & ", " & _
								m_oDataHistory(i).ANACDelay & ", " & _
								m_oDataHistory(i).ClientDelay & ", " & _
								m_oDataHistory(i).Docs30d & ", " & _
								m_oDataHistory(i).Closed30d & "]" )
				If i < 20 Then
					If m_oDataHistory(i).ANAC > max Then max = m_oDataHistory(i).ANAC
					If m_oDataHistory(i).Client > max Then max = m_oDataHistory(i).Client
					If m_oDataHistory(i).Docs30d > max Then max = m_oDataHistory(i).Docs30d
					If m_oDataHistory(i).Closed30d > max Then max = m_oDataHistory(i).Closed30d
				End If
			Next
			max = CInt(max/10 + 1)*10
		 %>
			]);
		var options = {
			title: 'Gráfico de Número de Processos',<%
			If m_bSend = True Then %>
			width: 980,<%
			End If %>
			height: 340,
			chartArea: {left:32,top:25,width:'80%',height:'70%'},
			hAxis: { baselineColor: 'none', gridlines: { color: 'transparent' }, format: 'd/MMM'},
			vAxis: { viewWindow: {max: <%=max %>} }
		};

		var chart_div = document.getElementById('chart1_<%=m_sObject %>');
		var chart = new google.visualization.LineChart(chart_div);
		<%
		If m_bSend = True Then %>
		// Wait for the chart to finish drawing before calling the getImageURI() method.
		google.visualization.events.addListener(chart, 'ready', function () {
			chart_div.innerHTML = '<img id="chart1_<%=m_sObject %>" src="' + chart.getImageURI() + '">';
			console.log(chart_div.innerHTML);<%
			If m_iGoal > 0 Then %>
			var form = document.getElementById('form_<%=m_sObject %>');
			form.image64_1.value = chart.getImageURI();
			chk1<%=m_sObject %> = true;
			<%
			End If %>
		});
		<%
		End If %>
		chart.draw(data, options);
	}

	function isImgSaved<%=m_sObject %>() {
		var http1 = new XMLHttpRequest();
		http1.open("HEAD", "http://sar/Public/img_<%=Year(Date()) %><%=Month(Date()) %><%=Day(Date()) %>_chart1_<%=m_sObject %>.png", false);
		http1.send();
		var ret1 = http1.status;
		var http2 = new XMLHttpRequest();
		http2.open("HEAD", "http://sar/Public/img_<%=Year(Date()) %><%=Month(Date()) %><%=Day(Date()) %>_chart2_<%=m_sObject %>.png", false);
		http2.send();
		var ret2 = http2.status;
		//alert("isImgSaved<%=m_sObject %>()=" + ret1 + "/" + ret2 );
		return ret1!=404 && ret2!=404;
	}

	<%
	If m_iGoal > 0 Then %>
	google.setOnLoadCallback(drawChart2<%=m_sObject %>);
	function drawChart2<%=m_sObject %>() {
		// Create and populate the data table.
		var data = new google.visualization.DataTable();
		data.addColumn('date', 'Date');
		data.addColumn('number', 'ANAC GOAL');
		data.addColumn('number', 'ANAC AVG');
		data.addColumn('number', 'ANAC MAX');
		data.addRows([
		<%
			first = True
			max = 10
			For i=m_nDataHistory-1 To 0 Step -1
				If Not first Then Response.Write("," & vbCrLf)
				first = False
				Response.Write("	[new Date(" & Year(m_oDataHistory(i).dtDate) & ", " & Month(m_oDataHistory(i).dtDate) - 1 & ", " & Day(m_oDataHistory(i).dtDate) & "), " & _
								 m_oDataHistory(i).Goal & ", " & _
								  m_oDataHistory(i).ANACAvg & ", " & _
								   m_oDataHistory(i).ANACMaxDays & "]")
				If i < 20 Then
					If m_oDataHistory(i).ANACMaxDays > max Then max = m_oDataHistory(i).ANACMaxDays
				End If
			Next
			max = CInt(max/10 + 1)*10
		 %>
			]);
		var options = {
			title: 'Gráfico de Tempos em Dias',<%
			If m_bSend = True Then %>
			width: 980,<%
			End If %>
			height: 340,
			chartArea: {left:32,top:25,width:'80%',height:'70%'},
			hAxis: { baselineColor: 'none', gridlines: { color: 'transparent' }, format: 'd/MMM'},
			vAxis: { viewWindow: {max: <%=max %>} }
		};
		var chart_div = document.getElementById('chart2_<%=m_sObject %>');
		var chart = new google.visualization.LineChart(chart_div);
		<%
		If m_bSend = True Then %>
		// Wait for the chart to finish drawing before calling the getImageURI() method.
		google.visualization.events.addListener(chart, 'ready', function () {
			chart_div.innerHTML = '<img id="chart2_<%=m_sObject %>" src="' + chart.getImageURI() + '">';
			console.log(chart_div.innerHTML);<%
			If m_iGoal > 0 Then %>
			var form = document.getElementById('form_<%=m_sObject %>');
			form.image64_2.value = chart.getImageURI();
			chk2<%=m_sObject %> = true;
			<%
			End If %>
		});
		<%
		End If %>
		chart.draw(data, options);
	}
	<%
	End If

	If m_bSend = True And m_iGoal > 0 Then %>
	var chkFiles<%=m_sObject %> = setInterval(function () { fChkFiles<%=m_sObject %>() }, 1000);
	function fChkFiles<%=m_sObject %>() {
		if( chk1<%=m_sObject %> == true && chk2<%=m_sObject %> == true ) {
			clearInterval(chkFiles<%=m_sObject %>);
			//alert("fChkFiles<%=m_sObject %>() -> submit");
			var form = document.getElementById('form_<%=m_sObject %>');
			form.submit();
		}
	}<%
	End If %>

	</script>
	<%
	End Sub


	Public Sub PrintHtml(htmlView)
	
	If htmlView = True Then %>
  <table border="0" width="100%" cellpadding="0" cellspacing="0">
    <tr>
	    <td width=26 height=30><img
			 src="../img/icons/glyphicons_214_resize_small.png" width=24 height=24
             border=0 onclick="hideChart<%=m_sObject %>();" id="divChart<%=m_sObject %>" alt="Inibe Gráfico."></td>
        <td colspan=4 align="left"><font size="5" face="Calibri"><%=m_sTitle %> - <%=m_strGTAR %> - RBAC <%=m_RBAC %></font></td>
	</tr>
    <tr>
	    <td width=26>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
        <td align="center"><div id="gauge_<%=m_sObject %>"></div></td>
        <td align="center"><div id="piechart_<%=m_sObject %>"></div></td>
        <td align="right" valign="top" width="86"><%
		If Alarm() = True Then
			Response.Write("<img src='../img/icons/glyphicons_207_remove_2.png' style='display:none;' " & _
							"onclick='hideAlarm" & m_sObject & "();' width=8 height=8 id='SplashAlarm1" & m_sObject & "'><br>" & _
							"<img src='../img/sirene.gif' style='display:none;' width=85 height=93 id='SplashAlarm2" & m_sObject & "'>")
		Else
			Response.Write("&nbsp;")
		End If
		%></td>
		<td width="100%">&nbsp; </td>
    </tr>
  </table><%
		If m_bSend = True And m_iGoal > 0 Then %>
  <form id="form_<%=m_sObject %>" action="http://sar/inet2/stats/ImgUpld.asp" method="POST" target="_blank"><%
		End If %>
  <table border="0" width="100%" id="SplashChart<%=m_sObject %>">
    <tr>
        <td><div id="chart1_<%=m_sObject %>"></div><%
		If m_bSend = True And m_iGoal > 0 Then %>
            <input type="hidden" name="image64_1">
            <input type="hidden" name="filename_1" value="chart1_<%=m_sObject %>"><%
		End If %>
        </td>
    </tr>
    <tr>
        <td><div id="chart2_<%=m_sObject %>"></div><%
		If m_bSend = True And m_iGoal > 0 Then %>
            <input type="hidden" name="image64_2">
            <input type="hidden" name="filename_2" value="chart2_<%=m_sObject %>"><%
		End If %>
		</td>
    </tr>
  </table><%
		If m_bSend = True And m_iGoal > 0 Then %>
  </form><%
		End If 

	ElseIf m_bSend = True And m_iGoal > 0 And m_strGTAR = "SAR" Then

		m_bodyMail = m_bodyMail & vbCrLf & _
"  <table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf & _
"    <tr>" & vbCrLf & _
"        <td align=""left""><font size=""5"" face=""Calibri"">&nbsp;&nbsp;&nbsp; " & m_sTitle & " - " & m_strGTAR & " - RBAC " & m_RBAC & "</font></td>" & vbCrLf & _
"	</tr>" & vbCrLf & _
"  </table>" & "<br>" & vbCrLf & _
"  <table border=""0"">" & vbCrLf & _
"    <tr>" & vbCrLf & _
"        <td><img src=""cid:img_" & Year(Date()) & Month(Date()) & Day(Date()) & "_chart1_" & m_sObject & ".png""></td>" & vbCrLf & _
"    </tr>" & vbCrLf & _
"    <tr>" & vbCrLf & _
"        <td><img src=""cid:img_" & Year(Date()) & Month(Date()) & Day(Date()) & "_chart2_" & m_sObject & ".png""></td>" & vbCrLf & _
"    </tr>" & vbCrLf & _
"  </table>"

	End If

	Dim PCodi
	Dim	BCodi
	Dim	OrgPCodi
	Dim	BaseAbrev
	Dim	SolDescr
	Dim	PesName
	Dim	StsDescr
	Dim sLink
	Dim i
	Dim first : first = True

	For i=0 To 9

		If m_oAIRData.TopTenProcSolic(i) <> "" Then

			querySQL = "SELECT S.P" & m_RBAC & "_CODI, O.ORG_NABREV, TS.TSOL_DESCR, Pes.PES_NGUERRA, " & _
					   "TST.TST_DESCR, O.ORGP_CODI, B.B" & m_RBAC & "_CODI" & _
					   "    FROM ( ( ( (A" & m_RBAC & "_TabSolic AS TS INNER JOIN " & _
					   "(Pessoal AS Pes INNER JOIN A" & m_RBAC & "_Solicitacoes AS S ON Pes.PES_CODI = S.PES_CODI) " & _
					   "ON TS.TSOL_CODI = S.TSOL_CODI) INNER JOIN A" & m_RBAC & "_Processos AS P " & _
					   "ON S.P" & m_RBAC & "_CODI = P.P" & m_RBAC & "_CODI) INNER JOIN A" & m_RBAC & "_Bases AS B " & _
					   "ON P.B" & m_RBAC & "_CODI = B.B" & m_RBAC & "_CODI) INNER JOIN Organizacao AS O " & _
					   "ON B.ORG_CODI = O.ORG_CODI) INNER JOIN A" & m_RBAC & "_TabStatus AS TST ON S.TST_CODI = TST.TST_CODI" & _
					   "    WHERE S.P" & m_RBAC & "_S" & m_RBAC & "='" & m_oAIRData.TopTenProcSolic(i) & "'"
			Set rsDiv = oDbFDH.getRecSetRd(querySQL)
			If Not rsDiv.Eof Then

				If first = True Then
				
					If htmlView = True Then %>
      <br><br>
      <table width=940 border=0 cellpadding=0 cellspacing=0>
         <tr>
            <td align=left>
              <fieldset>
                <legend><font size="3" face="Calibri" color=navy><b>Processos com maior atraso:</b></font></legend>
                <table width=860 border=0 cellpadding=1 cellspacing=0>
                   <tr bgcolor=blue align=center>
                        <td width=90>
                           <font size="2" face="Calibri" color=white><b>Processo</b></font></td>
                        <td width=190>
                           <font size="2" face="Calibri" color=white><b>Base</b></font></td>
                        <td width=220>
                           <font size="2" face="Calibri" color=white><b>Solicitação</b></font></td>
                        <td width=170>
                           <font size="2" face="Calibri" color=white><b>Status</b></font></td>
                        <td width=150>
                           <font size="2" face="Calibri" color=white><b>Analista</b></font></td>
                        <td width=40>
                           <font size="2" face="Calibri" color=white><b>Dias</b>&nbsp;</font></td>
                   </tr>
				<%
					ElseIf m_bSend = True And m_iGoal > 0 And m_strGTAR = "SAR" Then
						m_bodyMail = m_bodyMail & vbCrLf & _
"      <table width=940 border=0 cellpadding=0 cellspacing=0>" & vbCrLf & _
"         <tr>" & vbCrLf & _
"            <td align=left>" & vbCrLf & _
"              <fieldset>" & vbCrLf & _
"                <legend><font size=3 face=Calibri color=navy><b>Processos com maior atraso:</b></font></legend>" & vbCrLf & _
"                <table width=860 border=0 cellpadding=1 cellspacing=0>" & vbCrLf & _
"                   <tr bgcolor=blue align=center>" & vbCrLf & _
"                        <td width=90>" & vbCrLf & _
"                           <font size=""1"" face=""Calibri"" color=white><b>Processo</b></font></td>" & vbCrLf & _
"                        <td width=190>" & vbCrLf & _
"                           <font size=""1"" face=""Calibri"" color=white><b>Base</b></font></td>" & vbCrLf & _
"                        <td width=220>" & vbCrLf & _
"                           <font size=""1"" face=""Calibri"" color=white><b>Solicitação</b></font></td>" & vbCrLf & _
"                        <td width=170>" & vbCrLf & _
"                           <font size=""1"" face=""Calibri"" color=white><b>Status</b></font></td>" & vbCrLf & _
"                        <td width=150>" & vbCrLf & _
"                           <font size=""1"" face=""Calibri"" color=white><b>Analista</b></font></td>" & vbCrLf & _
"                        <td width=40>" & vbCrLf & _
"                           <font size=""1"" face=""Calibri"" color=white><b>Dias</b>&nbsp;</font></td>" & vbCrLf & _
"                   </tr>"
					End If
					first = false
				End If

				PCodi = rsDiv( "P" & m_RBAC & "_CODI" )
				BCodi = rsDiv( "B" & m_RBAC & "_CODI" )
				OrgPCodi = rsDiv( "ORGP_CODI" )
				BaseAbrev = rsDiv( "ORG_NABREV" )
				SolDescr = rsDiv( "TSOL_DESCR" )
				PesName = rsDiv( "PES_NGUERRA" )
				StsDescr = rsDiv( "TST_DESCR" )
				sLink = "http://sar/AvGeral/ControleProcessosMntProcCons.asp?Sec=" & m_RBAC & "&OrgPCodi=" & OrgPCodi & "&BCodi=" & BCodi & "&Letr=" & Left(BaseAbrev,1) & "&PCodi=" & PCodi & "&FlagAcesso=True"
				rsDiv.Close()
				Dim bgrColor
				If i Mod 2 = 0 Then
					bgrColor = "#E9E9E9"
				Else
					bgrColor = "#FFFFFF"
				End If
				If htmlView = True Then %>
				   <tr bgcolor=<%=bgrColor %>>
                        <td width=90 align=center>
                            <font class=SarText><b><a href="<%=sLink %>" target="_blank"><%=PCodi %></a></b></font></td>
                        <td width=190>
                            <font class=SarText><%=BaseAbrev %></font></td>
                        <td width=220>
                            <font class=SarText><%=SolDescr %></font></td>
                        <td width=170>
                            <font class=SarText><%=StsDescr %></font></td>
                        <td width=150>
                            <font class=SarText><%=PesName %></font></td>
                        <td width=40 align=right>
							<font class=SarText><%=m_oAIRData.TopTenDelay(i) %>&nbsp;</font></td>
                   </tr>
				<%
				ElseIf m_bSend = True And m_iGoal > 0 And m_strGTAR = "SAR" Then
					m_bodyMail = m_bodyMail & vbCrLf & _
"                   <tr bgcolor=" & bgrColor & ">" & vbCrLf & _
"                        <td width=90 align=center>" & vbCrLf & _
"                            <font size=""1"" face=""Calibri"">" & PCodi  & "</font></td>" & vbCrLf & _
"                        <td width=190>" & vbCrLf & _
"                            <font size=""1"" face=""Calibri"">" & BaseAbrev & "</font></td>" & vbCrLf & _
"                        <td width=220>" & vbCrLf & _
"                            <font size=""1"" face=""Calibri"">" & SolDescr & "</font></td>" & vbCrLf & _
"                        <td width=170>" & vbCrLf & _
"                            <font size=""1"" face=""Calibri"">" & StsDescr & "</font></td>" & vbCrLf & _
"                        <td width=150>" & vbCrLf & _
"                            <font size=""1"" face=""Calibri"">" & PesName & "</font></td>" & vbCrLf & _
"                        <td width=40 align=right>" & vbCrLf & _
"                            <font size=""1"" face=""Calibri"">" & m_oAIRData.TopTenDelay(i) & "&nbsp;</font></td>" & vbCrLf & _
"                   </tr>"
				End If

			End If

		Else
			Exit For
		End If

	Next

	If first = false Then
		If htmlView = True Then %>
                </table>
              </fieldset>
            </td>
         </tr>
      </table>

	<%

		ElseIf m_bSend = True And m_iGoal > 0 Then
			m_bodyMail = m_bodyMail & vbCrLf & _
"                </table>" & vbCrLf & _
"              </fieldset>" & vbCrLf & _
"            </td>" & vbCrLf & _
"         </tr>" & vbCrLf & _
"      </table>"
		End if

	End If

	End Sub

	Private Function calcHistory()

		Dim ret, dt, diff
		Dim iret : iret = 1
		Dim x : x = -1	
		
		For i=0 To m_nAIRCalc - 1
			m_oAIRCalc(i).oData.Open()
			ret = m_oAIRCalc(i).oData.FetchStart(-1)
			If ret <= 0 Then
				calcHistory = -1
				Exit Function
			End If
		Next

		Do While ret > 0

			For i=0 To m_nAIRCalc - 1

				dt = m_oAIRCalc(i).oData.dtDate

				' segue o primeiro (1)
				If i = 0 Then
					m_nDate = dt
					If x < MaxDays And DateDiff("m", m_nDate, date()) < 3 Then
						x = x + 1
						m_oDataHistory(x).Import( m_oAIRCalc(i).oData )
						ret = m_oAIRCalc(i).oData.FetchNext()
					Else
						Exit Do
					End If
				Else
					' diff = dt - m_nDate
					diff = DateDiff("d", m_nDate, dt)
					If diff = 0 Then ' está sincronizado

						' calcula
						m_oDataHistory(x).Add( m_oAIRCalc(i).oData )

						' next
						ret = m_oAIRCalc(i).oData.FetchNext()

					ElseIf diff > 0 Then ' aqui o cara está adiantado

						m_oAIRCalc(i).oData.FetchNext()

					End If

				End If

			Next
			
		Loop

		calcHistory = x + 1

	End Function

End Class

%>
