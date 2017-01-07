<%

'----------------------------------------------------------------
'
'	Class cWDCert145
'
'	Date: 27/07/2014
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------
'
Const NCert = 5

'----------------------------------------------------------------
'
'	Class cWDCert145
'

Class cWDCert145

	'Declarations
	Private m_oAIRStats(5)		' AIRStatistics
	Private m_oAIRRules(5)		' AIRStatistics Rules
	Private	i

	' Variables
	Private m_strGTAR			' GTAR
	Private m_nDate				' Date
	Private m_sTitle			' Nome do Objeto

	Private TSolCodi(5)
	Private TSolDescr(5)

	'Class Initialization
	Private Sub Class_Initialize()
		' TSOL_CODI
		TSolCodi(1) = TSOL_CODI005	' Certificação Inicial
		TSolCodi(2) = TSOL_CODI002	' Alteração Dados COM/EO
		TSolCodi(3) = TSOL_CODI006	' Aceitação de Lista de Capacidade
		TSolCodi(4) = TSOL_CODI008	' MNT Fora de Sede
		TSolCodi(5) = TSOL_CODI009	' Execução de Serviço Excepcional
		' TSOL_CODI Description
		TSolDescr(1) = "Cert Ini"
		TSolDescr(2) = "AltCOM/EO"
		TSolDescr(3) = "List Cap"
		TSolDescr(4) = "ForaSede"
		TSolDescr(5) = "SrvExcep"
	End Sub
	Public Default Function construct( gtar )
		m_sTitle = "Cert145"
		m_strGTAR = UCase(gtar)
		Dim dt : dt = Date()
		For i=1 To NCert

			Set m_oAIRStats(i)	= (new cAIRStatistics)( Array("145", m_strGTAR, TSolCodi(i)) )

			'                                            ANAC and Delay or Max or ANAC and 60d/ANAC or ANAC and 30d/ANAC
			'                                          Client and Delay or Max
			'                                        Delivery and Delay or Max or Deliv and 14d/Del or Deliv and 7d/Deliv
			Set m_oAIRRules(i)	= (new cAIRStatsRules)( Array( m_sTitle & TSolCodi(i) & Right(m_strGTAR,2), _
														   0,      8,      60,     0,        0,         0,        1, _
														   0,      4,      45,     0,        0,         0,        1, _
														   0,      2,      30,     0,        0,         0,        0, _
														   0,      1,      20,     0,        0,         0,        0, _
														   0,      0,      15,     0,        0,         0,        0, _
														   3,      5,     120,    -1,       -1,        -1,       -1, _
														   3,      3,      90,    -1,       -1,        -1,       -1, _
														   2,      2,      60,    -1,       -1,        -1,       -1, _
														   1,      1,      30,    -1,       -1,        -1,       -1, _
														   0,      0,      15,    -1,       -1,        -1,       -1, _
														   0,      5,      14,     0,        0,         0,        1, _
														   0,      3,      10,     0,        0,         0,        1, _
														   0,      1,       7,     0,        0,         0,        1, _
														   0,      0,       4,     0,        0,         0,        0, _
														   0,      0,       2,     0,        0,         0,        0 ) )
			m_oAIRStats(i).GetValues(dt)
		Next
		Calculate()
        Set construct = Me
    End Function

	'Terminate Class
	Private Sub Class_Terminate()
	End Sub


	Public Function Alarm()
		For i=1 To NCert
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
	google.load("visualization", "1", {packages:["corechart"]});
	google.setOnLoadCallback(drawPieChart<%=m_sTitle %>);
	function drawPieChart<%=m_sTitle %>() {
		var data = google.visualization.arrayToDataTable([
			['Label', 'Value']<%
		Dim nClient : nClient = 0
		Dim nDelivery : nDelivery = 0
		For i=1 To NCert
			Response.Write("," & vbCRLF)
			Response.Write("			['" & TSolDescr(i) & "', " & m_oAIRStats(i).oData.ANAC & "]")
			nClient = nClient + m_oAIRStats(i).oData.Client
			nDelivery = nDelivery + m_oAIRStats(i).oData.Delivery
		Next 
		%>,
			['OM', <%=nClient %>],
			['DISTR.', <%=nDelivery %>]
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
		For i=1 To NCert
			Response.Write("," & vbCRLF)
			Response.Write("			['" & TSolDescr(i) & "', 0]")
		Next 
		%>
			]);
		var options = {
			width: <%=NCert*135 %>, height: 120,
			redFrom: 85, redTo: 100,
			yellowFrom: 60, yellowTo: 85,
			minorTicks: 5
		};
		var chart = new google.visualization.Gauge(document.getElementById('gauge_<%=m_sTitle %>'));
		chart.draw(data, options);<%
		For i=1 To NCert %>
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
        <td colspan=4 align="left"><a href="/AvGeral/CtrlProcMntRelatPendBySolic.asp?Solic=<%=TSolCodi(1) %>&SDiv=<%=strGTAR %>&Sec=145"
		 style="text-decoration:none; color:Black;" target="_blank"><font
		 size="5" face="Calibri"><strong>Certificação e Outros da Intranet SAR - RBAC 145</strong></font></a></td>
	    <td width=<%=(26+4)*NCert %> height=30 nowrap><%
		For i=1 To NCert
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
		For i=1 To NCert
			m_oAIRRules(i).Rules(m_oAIRStats(i).oData)
		Next
	End Sub


End Class

%>
