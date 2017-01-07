<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage = 1252 %>
<% session.LCID = 1046 'BRASIL %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->

<%
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
' Windows Authentication
Dim ret : ret = oAAA.WinAuthenticate(False)
If ret < 0 Then
	oAAA.Print()
End If
Dim sdiv : sdiv = oAAA.AuthentWinUserSDiv

Dim Title : Title = "Intranet SAR"

 %>
<!-- #include virtual = "/inet2/header-bs.asp" -->

<style>
    	a:hover
    	{
    		color:write;
    		background-color: #CAEAFF;
    		text-decoration: none;
    	}
		.style54 {
			font-size: 10px;
			font-weight: bold;
		}
		.style55 
		{
			font-size: 10px
		}
			font-family: verdana;
			font-weight: bold;
		}
		.sRed 
		{
			font-family: Helvetica Neue, Helvetica, Arial, sans-serif;
			font-size: 18px;
			color: #FF0000;
		}
		.sBlue 
		{
			font-family: Helvetica Neue, Helvetica, Arial, sans-serif;
			font-size: 18px;
			color: #0000FF;
		}
		.sBlack
		{
			font-family: Helvetica Neue, Helvetica, Arial, sans-serif;
			font-size: 18px;
			color: #000000;
		}
		.sPink 
		{
			font-family: Helvetica Neue, Helvetica, Arial, sans-serif;
			font-size: 18px;
			color: #CC00CC;
		}
		.sPurple
		{
			font-family: Helvetica Neue, Helvetica, Arial, sans-serif;
			font-size: 18px;
			color: #660066;
		}
		.sGreen 
		{
			font-family: Helvetica Neue, Helvetica, Arial, sans-serif;
			font-size: 18px;
			color: #009900;
		}
		.linha
		{
			border-bottom: 1px dashed #c9cacb;
			margin: 7px 12px 7px 12px;
		}
</style>

<script language="JavaScript" type="text/javascript">

function getDocHeight(doc) {
	doc = doc || document;
	// from http://stackoverflow.com/questions/1145850/get-height-of-entire-document-with-javascript
	var body = doc.body, html = doc.documentElement;
	var height = Math.max( body.scrollHeight, body.offsetHeight, html.clientHeight, html.scrollHeight, html.offsetHeight );
	return height;
}

function setIframeHeight(id,max) {
	var ifrm = document.getElementById(id);
	var doc = ifrm.contentDocument? ifrm.contentDocument: ifrm.contentWindow.document;
	ifrm.style.visibility = 'hidden';
	ifrm.style.height = "10px"; // reset to minimal height in case going from longer to shorter doc
	var height = getDocHeight( doc );
	if (height > max) {
		height = max
		ifrm.style.scrolling = "auto"
	}
	ifrm.style.height = height + 5 + "px";
	ifrm.style.visibility = 'visible';
}

</script>

</head>

<body topmargin=0 leftmargin=0 background="images/imagem_index_claro.png">

<table border="0">
<tr>
    <td valign="top">
	  <div align="left">
		  
		<table border="0" width=100%>
			<tr>
				<td valign="top">
				<!-- #include file="Regulamentacao/ConsInternaIntranet.asp" -->
				</td>
				<td style="width: 12px;"> </td>
				<td valign="top" align="left"><br>
				<div class="thumbnail" style="height: 254px; width: 194px;">
					<table border="0" width="100%">
						<tr>
<td valign="top" width="84"><a href="https://sistemas.anac.gov.br/sip/login.php?sigla_orgao_sistema=ANAC&sigla_sistema=SEI" target="_blank"><img
 src="inet2/img/SEI.png" width="180" height="140" hspace="2" alt=""></a></td>
						</tr>
						<tr>
<td valign="top"><a href="https://sistemas.anac.gov.br/sip/login.php?sigla_orgao_sistema=ANAC&sigla_sistema=SEI" target="_blank"><div class="sBlack" align="center"><strong>SEI - Sistema Eletrônico de Informações</strong></div></a></td>
						</tr>
					</table>
				</div>
				</td>
			</tr>
		</table>

		<div class="linha" style="width: 824px;"></div>

		<table border="0" align="center">
			<tr>
				<td>
					<iframe marginwidth="2" marginheight="1" id="iFrameSearch"
					src="AvGeral/CtrlProcMntRelatSearch.asp" frameborder="0"
					width="840" onload="setIframeHeight(this.id,232)"></iframe>
				</td>
			</tr>
		</table>

		<div class="linha" style="width: 824px;"></div>

		<!-- Dynamic Alerts -->
		<%
		If Date() < CDate("05/10/2015") Then %>
		<table border="0">
			<tr>
				<td style="width: 9px;"> </td>
				<td style="width: 832px;">
					<div class="alert alert-danger fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<table><tr><td>
						<a href="http://intranet.anac.gov.br/comunicacao/acontece/2015/acontece2015_176.html"><img src="inet2/img/GEA-96x132.png" width="96" height="132" hspace="0" alt=""></a>
						</td>
						<td style="width: 12px;"> </td>
						<td>
						<strong>Notícia! Nova publicação da SAR aos operadores RBAC 121 e 135!!</strong><br><br>
						A GCVC/GGAC/SAR publicou o <b>Guia da Empresa Aérea - GEA</b>, para os operadores regidos pelo
						RBAC 121 (Linha Aérea) ou RBAC 135 (Táxi Aéreo)!!!<br><br>
						Acesse o Guia clicando <a href=http://intranet.anac.gov.br/comunicacao/acontece/2015/acontece2015_176.html target="_blank">aqui</a>.
						</td>
						</tr>
						</table>
					</div>
				</td>
			</tr>
		</table>
		<%
		End If %>

		<%
		If Date() < CDate("15/10/2015") Then %>
		<table border="0">
			<tr>
				<td style="width: 9px;"> </td>
				<td style="width: 832px;">
					<div class="alert alert-info fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Atenção!</strong> 
						A lista de aniversariantes do mês foi movida para a página de pessoal.<br>
						Para acessá-la, por favor utilize a opção "Pessoal" no topo do menu principal.<br>
					</div>
				</td>
			</tr>
		<%
		End If %>


		<%
		If date() < CDate("30/07/2016") Then %>
		<table border="0">
			<tr>
				<td style="width: 9px;"> </td>
				<td style="width: 832px;">
					<div class="alert alert-danger fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Informe!</strong> A GGAC emitiu o Boletim Informativo Interno de Aeronavegabilidade
						edição 008/2016, referente aos meses de Abr-Mai-Jun/2016.
						Acesse o Informativo clicando <a href=http://sar/AvGeral/Arquivos/Boletim_Interno_SAR_008-2016.pdf target="_blank">aqui</a>.
					</div>
				</td>
			</tr>
		</table>
		<%
		End If %>

		<%
		If Date() < CDate("30/07/2016") Then %>
		<table border="0">
			<tr>
				<td style="width: 9px;"> </td>
				<td style="width: 832px;">
					<div class="alert alert-info fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Atenção!</strong> 
						Foram emitidas cinco novas Instruções Técnicas Transitórias de Aeronavegabilidade (ITTA) pela SAR:<br>
						&nbsp;&nbsp;&nbsp;&nbsp;1. <b>ITTA 021-001/2016/GTAI</b> - Meios de cumprimento para Certificação de Organização de Produção;<br>
						&nbsp;&nbsp;&nbsp;&nbsp;2. <b>ITTA 183-006/2016/GTAS</b> - Orientações sobre comprovação de atribuição no CREA para PCFs;<br>
						&nbsp;&nbsp;&nbsp;&nbsp;3. <b>ITTA 119-012/2016/GCVC</b> - Orientações sobre a análise de processos de autorização ILS CAT II e III;<br>
						&nbsp;&nbsp;&nbsp;&nbsp;4. <b>ITTA 119-013/2016/GCVC</b> - Extensão de prazo para itens categoria “B” ou “C” da MEL.<br>
						&nbsp;&nbsp;&nbsp;&nbsp;5. <b>ITTA 091-014/2016/GCVC</b> - Emissão de NCIA – Alteração na numeração da NCIA.<br>
						Acesse as ITTAs em: <a href=Regulamentacao/ITTA.asp>http://SAR/Regulamentacao/ITTA.asp</a>
					</div>
				</td>
			</tr>
		<%
		End If %>

		<%
		If Date() < CDate("31/08/2016") Then %>
		<table border="0">
			<tr>
				<td style="width: 9px;"> </td>
				<td style="width: 832px;">
					<div class="alert alert-info fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Atenção!</strong> 
						Para substituir o número da credencial de INSPAC nos formulários deve-se utilizar o número do SIAPE, identificador único dos servidores públicos.<br>
						A <b>ITTA 091-014/2016/GCVC</b> - Emissão de NCIA – Alteração na numeração da NCIA, trata desta questão para as NCIA.<br>
						Informamos que os formulários abaixo que utilizam a identificação também já estão atualizados e disponíveis em http://sar/Regulamentacao/Formularios.asp:<br>
						a) F-100-34D: LISTA DE VERIFICAÇÃO PARA REALIZAÇÃO DE VISTORIA DE AERONAVE OU EMISSÃO DE RCA;<br>
						b) F-100-38B : LAUDO COMPLEMENTAR DE VISTORIA DE AERONAVE;<br>
						c) F-100-39A:  LAUDO DE VISTORIA;<br>
						d) F-100-40B: ETIQUETA PARA COLAGEM EM CADERNETA DA AERONAVE;<br>
						e) F-100-44A: NOTIFICAÇÃO DE CONDIÇÃO IRREGULAR DE AERONAVE;<br>
						f) F-900-44 : LAUDO DE AERONAVE OPERAÇÃO RVSM.<br>
						Acesse as ITTAs em: <a href=Regulamentacao/ITTA.asp>http://SAR/Regulamentacao/ITTA.asp</a>
					</div>
				</td>
			</tr>
		<%
		End If %>

		<%
		If Date() < CDate("01/01/2015") Then %>
		<table border="0">
				<td style="width: 9px;"> </td>
				<td style="width: 411px;">
					<div class="alert alert-danger fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Notícia!</strong> Foram emitidas duas <a href=Regulamentacao/ITTA.asp>Instruções
						Técnicas Transitórias: ITTA ..............</a>, sobre aprovação de blablabla.
					</div>
				</td>
				<!-- td style="width: 9px;"> </td>
				<td style="width: 272px;">
					<div class="alert fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Aviso!</strong> Se você ou melhorias para a Intranet SAR,
						informe aos <a href=http://sar/Gestores.asp>gestores dos subsistemas</a>.
					</div>
				</td>
				<td style="width: 9px;"> </td>
				<td style="width: 272px;">
					<div class="alert alert-danger fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Holy guacamole!</strong> Best check yo self, you're not looking too good.
					</div>
				</td -->
			</tr>
		</table>
		<%
		End If %>

		<%
		If date() < CDate("21/05/2016") Then %>
		<table border="0">
			<tr>
				<td style="width: 9px;"> </td>
				<td style="width: 832px;">
					<div class="alert alert-info fade in">
						<button type="button" class="close" data-dismiss="alert">&times;</button>
						<strong>Atenção!</strong> Foi publicada uma revisão da ITTA sobre EFB (ITTA n° 119-002/15/GCVC/GCAC/SAR) que alterou a seção
						6.3 do documento, retirando a orientação sobre o tamanho mínimo aceitável do display do PED. Ao longo do processo de aprovação do
						uso do PED para o operador, a avaliação do tamanho adequado do display será feita pelos inspetores da Superintendência de Padrões
						Operacionais (SPO).<br>
						Acesse as ITTAs em: <a href=Regulamentacao/ITTA.asp>http://SAR/Regulamentacao/ITTA.asp</a>
					</div>
				</td>
			</tr>

		</table>
		<%
		End If %>

		<table border="0" align="center">
			<tr>
				<td>
					<iframe marginwidth="2" marginheight="1" id="iFrameProcs"
					src="AvGeral/CtrlProcMntRelatUserPend.asp?SDiv=<%=sdiv %>" frameborder="0"
					width="840" onload="setIframeHeight(this.id,382)"></iframe>
				</td>
			</tr>
		</table>

		<div class="linha" style="width: 824px;"></div>

		<table>
		  <tr>

			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 84px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="80"><a href="AvGeral/AIR145.asp"><img src="inet2/img/Repair72x72.png" width="72" height="72" hspace="2" alt=""></div></a></td>
							<td valign="top"><a href="AvGeral/AIR145.asp"><div class="sBlue"><strong>Organizações de Manutenção<br>RBAC 145</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 84px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="80"><a href="../inet2/stats/AIRStats.asp"><img src="inet2/img/Statpic72x72.png" width="72" height="72" hspace="2" alt=""></div></a></td>
							<td valign="top"><a href="../inet2/stats/AIRStats.asp"><div class="sBlue"><strong>Estatísticas de atendimento aos Processos da GGAC</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 84px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="80"><a href="/BoletinsGGAC.html"><img src="inet2/img/boletim58x72.png" width="58" height="72" hspace="2" alt=""></div></a></td>
							<td valign="top"><a href="/BoletinsGGAC.html"><div class="sBlue"><strong>Boletins GGAC Informativos de Aeronavegabilidade</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

		  </tr>

		</table>

		<div class="linha" style="width: 824px;"></div>

		<table>
		  <tr>
			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 144px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="120"><a href="http://www2.anac.gov.br/publicacoes/Guia_Operador_Aeroagricola.html"><img src="inet2/img/GOA-100x132.png" width="100" height="132" hspace="0" alt=""></a></td>
							<td valign="top"><a href="http://www2.anac.gov.br/publicacoes/Guia_Operador_Aeroagricola.html"><div class="sBlue"><br><strong>GOA<br>Guia do Operador Aeroagrícola</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 144px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="120"><a href="http://www2.anac.gov.br/Publicacoes/Gea.html"><img src="inet2/img/GEA-96x132.png" width="96" height="132" hspace="0" alt=""></a></td>
							<td valign="top"><a href="http://www2.anac.gov.br/Publicacoes/Gea.html"><br><div class="sBlue"><strong>GEA<br>Guia da Empresa Aérea RBAC 121/135</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

		  </tr>

		</table>

		<div class="linha" style="width: 824px;"></div>

		<table>
		  <tr>
			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 84px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="80"><a href="inet2/gti.asp"><img src="inet2/img/IntranetSAR72x72.png" width="72" height="72" hspace="2" alt=""></a></td>
							<td valign="top"><a href="inet2/gti.asp"><div class="sBlack"><strong>Gestão de TI-SAR</strong> Estatísticas e outras informações.</div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 92px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="84"><a href="Anac/Arquivos/Desempenho_SAR.pdf" target="_blank"><img src="inet2/img/Stats120x80.png" width="120" height="80" hspace="2" alt=""></a></td>
							<td valign="top"><a href="http://intranet.anac.gov.br/fortalecimento_institucional/desempenho_institucional.html" target="_blank"><div class="sBlue"><strong>GTPA - Metas Institucionais &nbsp;&nbsp;&nbsp;Siga aqui!</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

			<td style="width: 8px;"> </td>

			<td>
				<div class="thumbnail" style="height: 92px; width: 272px;">
					<table border="0" width="100%">
						<tr>
							<td valign="top" width="84"><a href="http://compartilha-sar.anac.gov.br/gtgc/default.aspx" target="_blank"><img src="inet2/img/GC002-72x72.png" width="72" height="72" hspace="2" alt=""></a></td>
							<td valign="top"><a href="http://compartilha-sar.anac.gov.br/gtgc/default.aspx" target="_blank"><div class="sBlue"><strong>GTGC/SAR<BR>Gestão do Conhecimento</strong></div></a>
							</td>
						</tr>
					</table>
				</div>
			</td>

		  </tr>
		</table>


	  </div>
    </td>



  </tr>
</table>

<br><br><br><br>

<!-- #include virtual = "/inet2/trailer-bs.asp" -->


