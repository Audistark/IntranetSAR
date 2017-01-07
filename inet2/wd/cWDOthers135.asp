<%

'----------------------------------------------------------------
'
'	Class cWDOthers135
'
'	Date: 27/07/2014
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------
'
Const nOthers135 = 4
Const nOthersView135 = 3

'----------------------------------------------------------------
'
'	Class cWDOthers135
'

Class cWDOthers135

	'Declarations
	Private m_oAIRStats(4)		' AIRStatistics
	Private m_oAIRRules(3)		' AIRStatistics Rules
	Private m_oAIRData(3)		' AIRStatistics Data
	Private	i

	' Variables
	Private m_strGTAR			' GTAR
	Private m_nDate				' Date
	Private m_sTitle			' Nome do Objeto

	Private TSolCodi(4)
	Private TSolDescr(3)

	'Class Initialization
	Private Sub Class_Initialize()
		' TSOL_CODI
		TSolCodi(1) = TSOL_TXCODI007	' Autorização Especial
		TSolCodi(2) = TSOL_TXCODI008	' Dir Manutenção
		TSolCodi(3) = TSOL_TXCODI023	' Auditoria PTA
		TSolCodi(4) = TSOL_TXCODI006	' Auditoria Demanda
		' TSOL_CODI Description
		TSolDescr(1) = "Aut. Excep"
		TSolDescr(2) = "Cad DirMnt"
		TSolDescr(3) = "Auditorias"
	End Sub
	Public Default Function construct( gtar )
		m_sTitle = "Others135"
		m_strGTAR = UCase(gtar)
		Dim dt : dt = Date()
		For i=1 To nOthers135
			Set m_oAIRStats(i)	= (new cAIRStatistics)( Array("135", m_strGTAR, TSolCodi(i)) )
		Next

		'                                            ANAC and Delay or Max or ANAC and 60d/ANAC or ANAC and 30d/ANAC
		'                                          Client and Delay or Max
		'                                        Delivery and Delay or Max or Deliv and 14d/Del or Deliv and 7d/Deliv
		Set m_oAIRRules(1) = (new cAIRStatsRules)( Array( m_sTitle & Right(m_strGTAR,2), _
													   0,      3,      45,     0,        0,         0,        1, _
													   0,      0,      30,     0,        0,         0,        1, _
													   0,      0,      20,     0,        0,         0,        0, _
													   0,      0,      10,     0,        0,         0,        0, _
													   0,      0,       5,     0,        0,         0,        0, _
													   0,      5,     360,    -1,       -1,        -1,       -1, _
													   0,      4,     180,    -1,       -1,        -1,       -1, _
													   0,      3,      90,    -1,       -1,        -1,       -1, _
													   0,      2,      60,    -1,       -1,        -1,       -1, _
													   0,      0,      30,    -1,       -1,        -1,       -1, _
													   0,      2,      14,     0,        1,         0,        1, _
													   0,      1,       9,     0,        0,         0,        1, _
													   0,      0,       7,     0,        0,         0,        1, _
													   0,      0,       5,     0,        0,         0,        0, _
													   0,      0,       3,     0,        0,         0,        0 ) )
		Set m_oAIRRules(2) = (new cAIRStatsRules)( Array( m_sTitle & Right(m_strGTAR,2), _
													   0,      4,      60,     0,        1,         0,        1, _
													   0,      2,      45,     0,        0,         0,        1, _
													   0,      0,      30,     0,        0,         0,        1, _
													   0,      0,      15,     0,        0,         0,        0, _
													   0,      0,       5,     0,        0,         0,        0, _
													   0,      5,     360,    -1,       -1,        -1,       -1, _
													   0,      3,     180,    -1,       -1,        -1,       -1, _
													   0,      2,      90,    -1,       -1,        -1,       -1, _
													   0,      1,      60,    -1,       -1,        -1,       -1, _
													   0,      0,      30,    -1,       -1,        -1,       -1, _
													   0,      3,      14,     0,        1,         0,        1, _
													   0,      2,       9,     0,        1,         0,        1, _
													   0,      0,       7,     0,        0,         0,        1, _
													   0,      0,       5,     0,        0,         0,        0, _
													   0,      0,       3,     0,        0,         0,        0 ) )
		Set m_oAIRRules(3) = (new cAIRStatsRules)( Array( m_sTitle & Right(m_strGTAR,2), _
													   0,     10,      90,     0,        1,         0,        1, _
													   0,      5,      60,     0,        1,         0,        1, _
													   0,      0,      30,     0,        0,         0,        1, _
													   0,      0,      15,     0,        0,         0,        0, _
													   0,      0,       5,     0,        0,         0,        0, _
													   0,      8,     120,    -1,       -1,        -1,       -1, _
													   0,      5,      90,    -1,       -1,        -1,       -1, _
													   0,      3,      60,    -1,       -1,        -1,       -1, _
													   0,      1,      45,    -1,       -1,        -1,       -1, _
													   0,      0,      30,    -1,       -1,        -1,       -1, _
													   0,      3,      14,     0,        1,         0,        1, _
													   0,      2,       9,     0,        0,         0,        1, _
													   0,      0,       7,     0,        0,         0,        1, _
													   0,      0,       5,     0,        0,         0,        0, _
													   0,      0,       3,     0,        0,         0,        0 ) )
		For i=1 To nOthersView135
			m_oAIRStats(i).GetValues(dt)
			Set m_oAIRData(i) = new cAIRStatsData
		Next
		Calculate()
        Set construct = Me
    End Function


	'Terminate Class
	Private Sub Class_Terminate()
	End Sub


	Public Function Alarm()
		For i=1 To nOthersView135
			If m_oAIRRules(i).resDelayANAC = RES_CRITICO Or _
			   m_oAIRRules(i).resDelayClient = RES_CRITICO Or _
			   m_oAIRRules(i).resDelayDelivery = RES_CRITICO Then
				Alarm = True
				Exit For
			Else
				Alarm = False
			End If
		Next
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
   <%
	Dim nAnac : nAnac = 0
	Dim nClient : nClient = 0
	Dim nDelivery : nDelivery = 0
	%>
	google.load("visualization", "1", {packages:["corechart"]});
	google.setOnLoadCallback(drawPieChart<%=m_sTitle %>);
	function drawPieChart<%=m_sTitle %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value']<%
		For i=1 To nOthersView135
			Response.Write("," & vbCRLF)
			Response.Write("			['" & TSolDescr(i) & "', " & m_oAIRData(i).ANAC & "]")
			nAnac = nAnac + m_oAIRData(i).ANAC
			nClient = nClient + m_oAIRData(i).Client
			nDelivery = nDelivery + m_oAIRData(i).Delivery
		Next 
		%>
			]);
		var options = {
			backgroundColor: {fill:'white'},
			height: 120, width: 200,
			pieSliceText: 'value',
			enableInteractivity: true,
			chartArea: {left:0,top:6,width:'100%',height:'90%'}
		};
		var chart = new google.visualization.PieChart(document.getElementById('piechart_<%=m_sTitle %>'));
		chart.draw(data, options);
	}
	google.load("visualization", "1", {packages:["corechart"]});
	google.setOnLoadCallback(drawPieChart2<%=m_sTitle %>);
	function drawPieChart2<%=m_sTitle %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value'],
			['ANAC', <%=nAnac %>],
			['CLIENT', <%=nClient %>],
			['DISTR.', <%=nDelivery %>]
			]);
		var options = {
			backgroundColor: {fill:'white'},
			height: 120, width: 200,
			pieSliceText: 'value',
			colors: ['CornflowerBlue','Coral','LimeGreen'],
			enableInteractivity: true,
			chartArea: {left:0,top:6,width:'100%',height:'90%'}
		};
		var chart = new google.visualization.PieChart(document.getElementById('piechart2_<%=m_sTitle %>'));
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
			['Label', 'Value']<%
		For i=1 To nOthersView135
			Response.Write("," & vbCRLF)
			Response.Write("			['" & TSolDescr(i) & "', 0]")
		Next 
		%>
			]);
		var options = {
			width: <%=nOthersView135*135 %>, height: 120,
			redFrom: 85, redTo: 100,
			yellowFrom: 60, yellowTo: 85,
			minorTicks: 5
		};
		var chart = new google.visualization.Gauge(document.getElementById('gauge_<%=m_sTitle %>'));
		chart.draw(data, options);<%
		For i=1 To nOthersView135 %>
		setInterval(function() {
			var val=Math.floor((<%=m_oAIRRules(i).resDelayMax %>-0.50)*20 + 15*Math.random())
			if( val > 100 ) val = 100
			if( val <= 25 ) val = Math.floor(val*0.75)
			if( val < 0 ) val = 0
			data.setValue(<%=i-1 %>, 1, val);
			chart.draw(data, options);
			}, <%=i*200+5000 %>);
		<%
		Next %>
	}
  </script>
	<%
	End Sub


	Public Sub PrintCharts()
	End Sub
	

	Public Sub PrintHtml() %>

  <table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td>&nbsp; </td>
        <td colspan=5 align="left"><a href="/AvGeral/CtrlProcMntRelatPendBySolic.asp?Solic=<%=TSolCodi(1) %>&SDiv=<%=strGTAR %>&Sec=135"
		 style="text-decoration:none; color:Black;" target="_blank"><font
		 size="5" face="Calibri"><strong>Outros Processos de Táxi Aéreo - RBAC 135</strong></font></a></td>
	    <td width=<%=(26+4)*nOthersView135 %> height=30 nowrap><%
		For i=1 To nOthersView135
		%> <a href="javascript:showRules('<%=m_sTitle %><%=TSolCodi(i) %><%=Right(m_strGTAR,2) %>');"
		 title="Exibe regras de <%=m_sTitle %> - <%=TSolDescr(i) %>" style="text-decoration:none"><img
		 src="../img/icons/glyphicons_039_notes.png" width=20 height=25 border=0></a><%
		Next
		%></td>
    </tr>
    <tr>
	    <td width=26>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
        <td align="center"><div id="gauge_<%=m_sTitle %>"></div></td>
        <td align="center"><div id="piechart_<%=m_sTitle %>"></div></td>
        <td align="center"><div id="piechart2_<%=m_sTitle %>"></div></td>
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
		rows="6" cols="32" maxlength="200"><%=getNotes() %></textarea></td>
		<td>&nbsp; </td>
    </tr>
  </table>
	<%
	End Sub

	Private Sub Calculate()
		' Autorização Especial
		m_oAIRData(1).Import(m_oAIRStats(1).oData)
		m_oAIRRules(1).Rules(m_oAIRData(1))
		' Dir Manutenção
		m_oAIRData(2).Import(m_oAIRStats(2).oData)
		m_oAIRRules(2).Rules(m_oAIRData(2))
		' Auditorias
		m_oAIRData(3).Import(m_oAIRStats(3).oData)
		m_oAIRData(3).Add(m_oAIRStats(4).oData)
		m_oAIRRules(3).Rules(m_oAIRData(3))
	End Sub

End Class

%>
