<html>
<head>
<title>INTRANET - SAR</title>
  <style type="text/css">
    ul.menuLado
    {
        margin:0;
        padding:0;
    }
    ul.menuLado li
    {
        list-style:none
    }
    ul.menuLado li a
    {
        float:left;
	    width:132px;
	    font:12px bold Verdana, Arial, Helvetica, sans-serif;
	    font-weight:bold;
	    background:Navy;
	    color:white;
	    text-align:left;
	    padding:0.2em 0.2em 0em 0;
	    text-decoration:none;
    }
    ul.menuLado li :hover{
		background:white;
		color:Navy;
		border-color:white;
	}
  </style>
  <script language="JavaScript" type="text/javascript">
  	function HideSide() {
  		var fs = window.top.document.getElementsByTagName("frameset");
  		fs[1].cols = "10,*"
  	}
  	function ShowSide() {
  		var fs = window.top.document.getElementsByTagName("frameset");
  		if (fs[1].cols == "10,*" || fs[1].cols == "8,*")
			fs[1].cols = "132,*"
  	}
  </script>
</head>
<body bgcolor=Navy text="#000000" topmargin="0" leftmargin="4" onmouseover="JavaScript:ShowSide();">
<div align="center">
	<ul class="menuLado">
		<li>&nbsp;&nbsp;&nbsp;</li>
		<li><a href="Pessoal/Pessoal.asp" target="mainFrame">Pessoal</a></li>
		<li><a href="processo.asp" target="mainFrame">Processos</a></li>
		<li><a href="Engenharia/Engenharia.asp" target="mainFrame">Engenharia</a></li>
		<li><a href="Programas/Programas.asp" target="mainFrame">Programas</a></li>
		<li><a href="AvGeral/AuditoriaInspecao.asp" target="mainFrame">Auditoria/Inspeção</a></li>
		<li><a href="Regulamentacao/Regulamentacao.asp" target="mainFrame">Processo Normativo</a></li>
		<li><a href="secretaria1.asp" target="mainFrame">Documentação</a></li>
		<li><a href="Financeiro/menu.asp" target="mainFrame">Financeiros</a></li>
		<li><a href="AvGeral/AvGeral.asp" target="mainFrame">Aeronavegabilidade</a></li>
		<li>&nbsp;&nbsp;&nbsp;</li>
		<li><a href="http://frequencia.anac.gov.br" target="_blank">Registro Frequência</a></li>
		<li><a href="https://correio.anac.gov.br/owa" target="_blank">Correio ANAC</a></li>
		<li><a href="https://servicosti.anac.gov.br" target="_blank">Serviços TI</a></li>
		<li><a href="mailto:suporteti@anac.gov.br" target="_blank">Help Desk STI</a></li>
		<li><a href="Changelog.asp" target="mainFrame">Changelog Intranet</a></li>
		<li><a href="Gestores.asp" target="mainFrame">Gestores TI SAR</a></li>
		<li><a href="http://intranet.anac.gov.br/sso/" target="_blank">Intranet - SSO</a></li>
		<li>&nbsp;&nbsp;&nbsp;</li>
		<li><a href="Senha.asp?Arq=_Mnt/Manutencao.asp" target="mainFrame">Manutenção (Login)</a></li>
		<li>&nbsp;&nbsp;&nbsp;</li>
		<li>&nbsp;&nbsp;&nbsp;</li>
	</ul>
</div>
<div>
	<ul style="list-style:none">
		<li><a href="JavaScript:HideSide();"><img
		 width="23" height="23" border="0"
		 src="inet2/img/icons/glyphicons_335_pushpin.png"></a></li>
	</ul>
</div>
</body>
</html>
