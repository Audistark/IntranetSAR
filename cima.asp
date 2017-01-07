<html>
<head> 
<title>INTRANET - SAR</title>
<link href="Images/globo_anac.ico" rel="shortcut icon" type="image/x-icon"/>
 <style>
    ul.menucima
    {
		margin:0;
		padding:0;
    }
    ul.menucima li
    {
		list-style:none;
		display:inline;
    }
    ul.menucima li a
    {
		float:left;
		border:1px solid #303EA5;
		width:150px;
		font:12px bold Verdana, Arial, Helvetica, sans-serif;
		font-weight:bold;
		background:#303EA5;
		color:#ADD6F3;
		text-align:center;
		text-decoration:none;
    }
    ul.menucima li :hover{
		background:#ADD6F3;
		color:#303EA5;
		border-color:#ADD6F3;
	}
 </style>

 <script language="JavaScript" type="text/javascript">
 	function HideTop() {
 		var fs = window.top.document.getElementsByTagName("frameset");
 		fs[0].rows = "10,*"
 	}
 	function ShowTop() {
 		var fs = window.top.document.getElementsByTagName("frameset");
 		if (fs[0].rows == "10,*" || fs[0].rows == "8,*")
			fs[0].rows = "114,*"
 	}
 </script>

</head>
  
<body bgcolor="#303ea5" text="#000000" topmargin="0" leftmargin="0" onmouseover="JavaScript:ShowTop();">
  <img src="Images/barra-anac.jpg" width="780" height="91" border="0" usemap="#Map">
  <ul class="menucima">
	<li><a href="/" target="_top">Home</a></li>
	<li><a href="Pessoal/Pessoal.asp" target="mainFrame">Usuários</a></li>
	<li><a href="linkutil.asp" target="mainFrame">Links Úteis</a></li>
  </ul>
  <map name="Map">
    <area shape="poly" coords="0,52,0,52,32,51,32,77,109,78,111,51,139,53,139,16,-1,16" href="/" target="_top">
    <area shape="rect" coords="694,-1,776,88" href="http://www.anac.gov.br/" target="_blank">
    <area shape="rect" coords="640,80,680,88" href="JavaScript:HideTop();">
  </map>
</body>
</html>
