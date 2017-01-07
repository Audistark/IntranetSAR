<%

'----------------------------------------------------------------
'
'	Class cWDManuals145
'
'	Date: 27/07/2014
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------
'

'----------------------------------------------------------------
'
'	Class cWDManuals145
'

Class cWDManuals145

	'Declarations
	Private m_oAIRStats(4)		' AIRStatistics
	Private m_oAIRData			' AIRStatistics Data
	Private m_oAIRRules			' AIRStatistics Rules

	' Variables
	Private m_strGTAR			' GTAR
	Private m_nDate				' Date
	Private m_sTitle			' Nome do Objeto

	'Class Initialization
	Private Sub Class_Initialize()
	End Sub
	Public Default Function construct( gtar )
		m_sTitle = "Manuals145"
		m_strGTAR = UCase(gtar)
		Set m_oAIRStats(1) = (new cAIRStatistics)( Array("145", m_strGTAR, TSOL_CODI004) )
		Set m_oAIRStats(2) = (new cAIRStatistics)( Array("145", m_strGTAR, TSOL_CODI012) )
		Set m_oAIRStats(3) = (new cAIRStatistics)( Array("145", m_strGTAR, TSOL_CODI016) )
		Set m_oAIRStats(4) = (new cAIRStatistics)( Array("145", m_strGTAR, TSOL_CODI017) )
		'                                            ANAC and Delay or Max or ANAC and 60d/ANAC or ANAC and 30d/ANAC
		'                                          Client and Delay or Max
		'                                        Delivery and Delay or Max or Deliv and 14d/Del or Deliv and 7d/DelivManuals135	0;3;90;0;0,5;0;0,8;0;2;60;0;0;0;0,6;0;1;30;0;0;0;0,5;0;0;20;0;0;0;0,3;0;0;15;0;0;0;0,1;	5;5;150;-1;-1;-1;-1;4;4;120;-1;-1;-1;-1;3;3;90;-1;-1;-1;-1;2;2;60;-1;-1;-1;-1;1;0;45;-1;-1;-1;-1;	0;8;45;0;0,5;0;0,8;0;6;30;0;0;0;0,6;0;4;15;0;0;0;0,4;0;0;15;0;0;0;0,2;0;0;15;0;0;0;0,1;
		Set m_oAIRRules = (new cAIRStatsRules)( Array( m_sTitle & Right(m_strGTAR,2), _
													   0,      3,      90,     0,        1,         0,        1, _
													   0,      2,      60,     0,        0,         0,        1, _
													   0,      1,      30,     0,        0,         0,        0, _
													   0,      0,      20,     0,        0,         0,        0, _
													   0,      0,      15,     0,        0,         0,        0, _
													   5,      5,     150,    -1,       -1,        -1,       -1, _
													   4,      4,     120,    -1,       -1,        -1,       -1, _
													   3,      3,      90,    -1,       -1,        -1,       -1, _
													   2,      2,      60,    -1,       -1,        -1,       -1, _
													   1,      0,      45,    -1,       -1,        -1,       -1, _
													   0,     19,      45,     0,        1,         0,        1, _
													   0,     12,      30,     0,        1,         0,        1, _
													   0,      9,      15,     0,        1,         0,        1, _
													   0,      3,      10,     0,        0,         0,        1, _
													   0,      0,       7,     0,        0,         0,        0 ) )
		Set m_oAIRData = new cAIRStatsData
		Dim dt : dt = Date()
		m_oAIRStats(1).GetValues(dt)
		m_oAIRStats(2).GetValues(dt)
		m_oAIRStats(3).GetValues(dt)
		m_oAIRStats(4).GetValues(dt)
		Calculate()
        Set construct = Me
    End Function

	'Terminate Class
	Private Sub Class_Terminate()
	End Sub

	Public Function Alarm()
		If m_oAIRRules.resDelayANAC = RES_CRITICO Or _
		   m_oAIRRules.resDelayClient = RES_CRITICO Or _
		   m_oAIRRules.resDelayDelivery = RES_CRITICO Then
			Alarm = True
		Else
			Alarm = False
		End If
	End Function

	Private Function getNotes()
		getNotes = Request.Cookies("Note" & m_sTitle)
	End function

	Public Sub JavaScripts() %>
	<!-- #include virtual = "/inet2/wd/hWDIncludes.asp" -->
	<%
	End Sub

	Public Sub PrintPieChart() %>
  <script type="text/javascript">
	google.load("visualization", "1", {packages:["corechart"]});
	google.setOnLoadCallback(drawPieChart<%=m_sTitle %>);
	function drawPieChart<%=m_sTitle %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value'],
			['ANAC', <%=m_oAIRData.ANAC %>],
			['OM', <%=m_oAIRData.Client %>],
			['DISTR.', <%=m_oAIRData.Delivery %>]
			]);
		var options = {
			backgroundColor: {fill:'white'},
			height: 120, width: 200,
			pieSliceText: 'value',
			enableInteractivity: true,
			colors: ['CornflowerBlue','Coral','LimeGreen'],
			chartArea: {left:0,top:6,width:'100%',height:'90%'}
		};
		var chart = new google.visualization.PieChart(document.getElementById('piechart_<%=m_sTitle %>'));
		chart.draw(data, options);
      }
  </script>
	<%
	End Sub


	Public Sub PrintGauge() %>
  <script type="text/javascript">
	google.load('visualization', '1', { packages: ['gauge'] });
	google.setOnLoadCallback(drawGauges<%=m_sTitle %>);
	function drawGauges<%=m_sTitle %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value'],
			['ANAC', 0],
			['OM', 0],
			['DISTR.', 0]
			]);
		var options = {
			width: 400, height: 120,
			redFrom: 85, redTo: 100,
			yellowFrom: 60, yellowTo: 85,
			minorTicks: 5
		};
		var chart = new google.visualization.Gauge(document.getElementById('gauge_<%=m_sTitle %>'));
		chart.draw(data, options);
		setInterval(function() {
			var val=Math.floor((<%=m_oAIRRules.resDelayANAC %>-0.50)*20 + 15*Math.random())
			if( val > 100 ) val = 100
			if( val <= 25 ) val = Math.floor(val*0.75)
			if( val < 0 ) val = 0
			data.setValue(0, 1, val);
			chart.draw(data, options);
			}, 5000);
		setInterval(function() {
			var val=Math.floor((<%=m_oAIRRules.resDelayClient %>-0.50)*20 + 15*Math.random())
			if( val > 100 ) val = 100
			if( val <= 25 ) val = Math.floor(val*0.75)
			if( val < 0 ) val = 0
			data.setValue(1, 1, val);
			chart.draw(data, options);
			}, 5500);
		setInterval(function() {
			var val=Math.floor((<%=m_oAIRRules.resDelayDelivery %>-0.50)*20 + 15*Math.random())
			if( val > 100 ) val = 100
			if( val <= 25 ) val = Math.floor(val*0.75)
			if( val < 0 ) val = 0
			data.setValue(2, 1, val);
			chart.draw(data, options);
			}, 5800);
	}
  </script>
	<%
	End Sub

	Public Sub PrintCharts()
 %>
  <script type="text/javascript">
	google.load('visualization', '1', { packages: ['corechart'] });
   	google.setOnLoadCallback(drawChart<%=m_sTitle %>);
    	function drawChart<%=m_sTitle %>() {
		// Create and populate the data table.
		var data = new google.visualization.DataTable();
		data.addColumn('date', 'Date');
		data.addColumn('number', 'Total Procs');
		data.addColumn('number', 'ANAC');
		data.addColumn('number', 'OM');
		data.addColumn('number', 'ANAC>30d');
		data.addColumn('number', 'OM Delayed');
		data.addColumn('number', '30d Closed');
		data.addColumn('number', '30d Docs');
		data.addRows([
<%
		Dim first : first = True
		m_oAIRStats(1).oData.Open()
		m_oAIRStats(2).oData.Open()
		m_oAIRStats(3).oData.Open()
		m_oAIRStats(4).oData.Open()
		If m_oAIRStats(1).oData.FetchStart() > 0 And m_oAIRStats(2).oData.FetchStart() > 0 And _
		   m_oAIRStats(3).oData.FetchStart() > 0 And m_oAIRStats(4).oData.FetchStart() > 0 Then
			Do
				m_nDate = m_oAIRStats(1).oData.dtDate
				Do While m_nDate > m_oAIRStats(2).oData.dtDate
					If m_oAIRStats(2).oData.FetchNext() < 1 Then
						Exit Do
					End If
				Loop
				Do While m_nDate < m_oAIRStats(2).oData.dtDate
					If m_oAIRStats(1).oData.FetchNext() < 1 Then
						Exit Do
					End If
				Loop
				m_nDate = m_oAIRStats(2).oData.dtDate
				Do While m_nDate > m_oAIRStats(3).oData.dtDate
					If m_oAIRStats(3).oData.FetchNext() < 1 Then
						Exit Do
					End If
				Loop
				Do While m_nDate < m_oAIRStats(3).oData.dtDate
					If m_oAIRStats(2).oData.FetchNext() < 1 Then
						Exit Do
					End If
				Loop
				m_nDate = m_oAIRStats(3).oData.dtDate
				Do While m_nDate > m_oAIRStats(4).oData.dtDate
					If m_oAIRStats(4).oData.FetchNext() < 1 Then
						Exit Do
					End If
				Loop
				Do While m_nDate < m_oAIRStats(4).oData.dtDate
					If m_oAIRStats(3).oData.FetchNext() < 1 Then
						Exit Do
					End If
				Loop
				If m_nDate = m_oAIRStats(4).oData.dtDate And _
				   DateDiff("m", m_nDate, Date) < 6 And _
				   Weekday(m_nDate) <> vbSunday And Weekday(m_nDate) <> vbSaturday Then ' sábado e domingo não
					Call GetValues()
					If Not first Then Response.Write("," & vbCrLf)
					first = False
					Response.Write("			[new Date(" & Year(m_nDate) & ", " & Month(m_nDate) - 1 & ", " & Day(m_nDate) & "), " & _
									m_oAIRData.Sum & ", " & _
									 m_oAIRData.ANAC & ", " & _
									  m_oAIRData.Client & ", " & _
									   m_oAIRData.ANAC30d & ", " & _
									    m_oAIRData.ClientDelay & ", " & _
									     m_oAIRData.Closed30d & ", " & _
										  m_oAIRData.Docs30d & "]")
				End If
			Loop While m_oAIRStats(1).oData.FetchNext() > 0 And m_oAIRStats(2).oData.FetchNext() > 0 And _
					    m_oAIRStats(3).oData.FetchNext() > 0 And m_oAIRStats(4).oData.FetchNext() > 0
			m_oAIRStats(1).oData.Close()
			m_oAIRStats(2).oData.Close()
			m_oAIRStats(3).oData.Close()
			m_oAIRStats(4).oData.Close()
		End If
		 %>
			]);
		var options = {
			height: 340,
			chartArea: {left:32,top:25,width:'82%',height:'75%'},
			hAxis: { baselineColor: 'none', gridlines: { color: 'transparent' }, format: 'd/MMM'}
    	};
    	var chart = new google.visualization.LineChart(document.getElementById('chart_<%=m_sTitle %>'));
    	chart.draw(data, options);
    }
  </script>
	<%
	End Sub


	Public Sub PrintHtml() %>

  <table border="0" cellpadding="0" cellspacing="0">
    <tr>
	    <td width=26 height=30><img
			 src="../img/icons/glyphicons_214_resize_small.png" width=24 height=24
             border=0 onclick="hideChart<%=m_sTitle %>();" id="divChart<%=m_sTitle %>" alt="Inibe Gráfico."></td>
        <td colspan=4 align="left"><a href="/AvGeral/CtrlProcMntRelatPendBySolic.asp?Solic=<%=TSOL_CODI004 %>&SDiv=<%=strGTAR %>&Sec=145"
		 style="text-decoration:none; color:Black;" target="_blank"><font
         size="5" face="Calibri"><strong>Análise de Manuais (MOM/MCQ/PT/Supl./SGSO) - RBAC 145</strong></font></a></td>
	    <td width=26 height=30><img
			 src="../img/icons/glyphicons_039_notes.png" width=20 height=25
             border=0 onclick="showRules('<%=m_sTitle %><%=Right(m_strGTAR,2) %>');" alt="Exibe regras."></td>
    </tr>
    <tr>
        <td>&nbsp; </td>
        <td align="center"><div id="gauge_<%=m_sTitle %>"></div></td>
        <td align="center"><div id="piechart_<%=m_sTitle %>"></div></td>
        <td align="right" valign="top" width="86"><%
		If Alarm() = True Then
			Response.Write("<img src='../img/icons/glyphicons_207_remove_2.png' style='display:none;' " & _
							"onclick='hideAlarm" & m_sTitle & "();' width=8 height=8 id='SplashAlarm1" & m_sTitle & "'><br>" & _
							"<img src='../img/sirene.gif' style='display:none;' width=85 height=93 id='SplashAlarm2" & m_sTitle & "'>")
		Else
			Response.Write("&nbsp;")
		End If
		%></td>
		<td align="center" valign="top"><textarea id="Note<%=m_sTitle %>" OnKeyUp="setNote<%=m_sTitle %>();" style="border: none; overflow:hidden"
		rows="6" cols="38" maxlength="200"><%=getNotes() %></textarea></td>
		<td>&nbsp; </td>
    </tr>
  </table>

  <table border="0" width="100%">
    <tr id="SplashChart<%=m_sTitle %>">
        <td><div id="chart_<%=m_sTitle %>"></div></td>
    </tr>
  </table>

	<%
	End Sub

	Private Sub GetValues()
		m_oAIRData.Import(m_oAIRStats(1).oData)
		m_oAIRData.Add(m_oAIRStats(2).oData)
		m_oAIRData.Add(m_oAIRStats(3).oData)
		m_oAIRData.Add(m_oAIRStats(4).oData)
	End Sub

	Private Sub Calculate()
		Call GetValues()
		m_oAIRRules.Rules(m_oAIRData)
	End Sub

End Class

%>
