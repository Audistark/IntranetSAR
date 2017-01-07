
  <script type="text/javascript">

  	function showChart<%=m_sTitle %>() {
  		myTimer<%=m_sTitle %> = setInterval(function () { hideChart<%=m_sTitle %>() }, 30000);
  		// show SplashChart<%=m_sTitle %>
  		var div = document.getElementById("SplashChart<%=m_sTitle %>");
  		div.style.display = "inline";
		document.getElementById("divChart<%=m_sTitle %>").src = "../img/icons/glyphicons_214_resize_small.png";
		document.getElementById("divChart<%=m_sTitle %>").title = "Inibe Gráfico.";
		document.getElementById("divChart<%=m_sTitle %>").onclick = function () { hideChart<%=m_sTitle %>(); }
		document.getElementById("divChart<%=m_sTitle %>").width = 24;
		document.getElementById("divChart<%=m_sTitle %>").height = 24;
	}
	function hideChart<%=m_sTitle %>() {
		clearInterval(myTimer<%=m_sTitle %>);
		// hide SplashChart<%=m_sTitle %>
		var div = document.getElementById("SplashChart<%=m_sTitle %>");
		div.style.display = "none";
		document.getElementById("divChart<%=m_sTitle %>").src = "../img/icons/glyphicons_040_stats.png";
		document.getElementById("divChart<%=m_sTitle %>").title = "Mostra Gráfico.";
		document.getElementById("divChart<%=m_sTitle %>").onclick = function () { showChart<%=m_sTitle %>(); }
		document.getElementById("divChart<%=m_sTitle %>").width = 26;
		document.getElementById("divChart<%=m_sTitle %>").height = 25;
	}
	var myTimer<%=m_sTitle %> = setInterval(function () { hideChart<%=m_sTitle %>() }, 20000);
	function enableAlarm<%=m_sTitle %>() {
		clearInterval(myTimerAlert<%=m_sTitle %>);
		// show SplashAlarm<%=m_sTitle %>
		var div = document.getElementById("SplashAlarm1<%=m_sTitle %>");
		div.style.display = "inline";
		div = document.getElementById("SplashAlarm2<%=m_sTitle %>");
		div.style.display = "inline";
	}
	var myTimerAlert<%=m_sTitle %> = setInterval(function () { enableAlarm<%=m_sTitle %>() }, 8000);
	function hideAlarm<%=m_sTitle %>() {
		// hide SplashAlarm<%=m_sTitle %>
		var div = document.getElementById("SplashAlarm1<%=m_sTitle %>");
		div.style.display = "none";
		var div = document.getElementById("SplashAlarm2<%=m_sTitle %>");
		div.style.display = "none";
	}
	function setNote<%=m_sTitle %>() {
		var edt = document.getElementById("Note<%=m_sTitle %>").value;
		var d = new Date();
		var u = new Date();
		u.setDate(d.getDate() + 20000);
		document.cookie = "Note<%=m_sTitle %>=" + escape(edt) +"; expires=" + u.toUTCString();
	}
	function showRules(tag) {
		my_window = window.open("hWDRulesCfg.asp?tag=" + tag, "1372619272713", "width=740, height=480, toolbar=0, menubar=0, status=1, resizable=1");
	}

  </script>
