
  <script type="text/javascript">

  	function showChart<%=m_sObject %>() {
  		myTimer<%=m_sObject %> = setInterval(function () { hideChart<%=m_sObject %>() }, 600000);
  		// show SplashChart<%=m_sObject %>
  		var div = document.getElementById("SplashChart<%=m_sObject %>");
  		div.style.display = "inline";
		document.getElementById("divChart<%=m_sObject %>").src = "../img/icons/glyphicons_214_resize_small.png";
		document.getElementById("divChart<%=m_sObject %>").title = "Inibe Gráfico.";
		document.getElementById("divChart<%=m_sObject %>").onclick = function () { hideChart<%=m_sObject %>(); }
		document.getElementById("divChart<%=m_sObject %>").width = 24;
		document.getElementById("divChart<%=m_sObject %>").height = 24;
	}
	function hideChart<%=m_sObject %>() {
		clearInterval(myTimer<%=m_sObject %>);
		// hide SplashChart<%=m_sObject %>
		var div = document.getElementById("SplashChart<%=m_sObject %>");
		div.style.display = "none";
		document.getElementById("divChart<%=m_sObject %>").src = "../img/icons/glyphicons_040_stats.png";
		document.getElementById("divChart<%=m_sObject %>").title = "Mostra Gráfico.";
		document.getElementById("divChart<%=m_sObject %>").onclick = function () { showChart<%=m_sObject %>(); }
		document.getElementById("divChart<%=m_sObject %>").width = 26;
		document.getElementById("divChart<%=m_sObject %>").height = 25;
	}
	var myTimer<%=m_sObject %> = setInterval(function () { hideChart<%=m_sObject %>() }, 600000);
	function enableAlarm<%=m_sObject %>() {
		clearInterval(myTimerAlert<%=m_sObject %>);
		// show SplashAlarm<%=m_sObject %>
		var div = document.getElementById("SplashAlarm1<%=m_sObject %>");
		div.style.display = "inline";
		div = document.getElementById("SplashAlarm2<%=m_sObject %>");
		div.style.display = "inline";
	}
	var myTimerAlert<%=m_sObject %> = setInterval(function () { enableAlarm<%=m_sObject %>() }, 8000);
	function hideAlarm<%=m_sObject %>() {
		// hide SplashAlarm<%=m_sObject %>
		var div = document.getElementById("SplashAlarm1<%=m_sObject %>");
		div.style.display = "none";
		var div = document.getElementById("SplashAlarm2<%=m_sObject %>");
		div.style.display = "none";
	}
	function setNote<%=m_sObject %>() {
		var edt = document.getElementById("Note<%=m_sObject %>").value;
		var d = new Date();
		var u = new Date();
		u.setDate(d.getDate() + 20000);
		document.cookie = "Note<%=m_sObject %>=" + escape(edt) +"; expires=" + u.toUTCString();
	}

  </script>
