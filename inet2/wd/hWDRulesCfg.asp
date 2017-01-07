<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.CodePage = 1252 %>

<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->

<%

	' Init AAA - Authentication, Authorization and Accounting
	Dim oAAA : Set oAAA = new cAAA
	Dim ret : ret = oAAA.WinAuthenticate(True)
	If ret < 0 Then
		Response.Status = "403 Forbidden"
		Response.End
	End If

	' 
	Dim tag : tag = Request.QueryString("tag")
	Dim cls : cls = Request.QueryString("close")
	Dim m_rVal(3,5,7)
	Dim oDbFDH, rsDiv, querySQL

	' Regulation
	Dim sec : sec = ""
	Dim pos : pos = InStr(tag,"91")
	pos = InStr(tag,"91")
	If pos > 0 Then
		sec = "91"
	Else
	pos = InStr(tag,"121")
	If pos > 0 Then
		sec = "121"
	Else
	pos = InStr(tag,"135")
	If pos > 0 Then
		sec = "135"
	Else
	pos = InStr(tag,"145")
	If pos > 0 Then
		sec = "145"
	End If
	End If
	End If
	End If

	' Gtar
	Dim gtar : gtar = "GTAR-" & Right(tag,2)

	' Só MASTER, GESTORES e LÍDERES podem alterar
	Dim readonly : readonly = ""
	' access control
	If oAAA.AuthorWinMaster() <> True And _
	   oAAA.AuthorWinMasterSec(sec) <> True And _
	   ( oAAA.AuthorWinLiderSec(sec) <> True Or oAAA.AuthentWinUserSDiv() <> gtar ) Then
		readonly = " readonly"
		If cls = "true" Then %>
			<script type="text/javascript">
				window.close();		
			</script>
			<%
			Response.End
		End If
	End If

	' DB Connection
	Set oDbFDH = (new cDBAccess)("FDH")
	querySQL = "SELECT * FROM WDRules WHERE WDRules_Id = '" & tag & "'"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If Not rsDiv.Eof Then
		For i=0 To 2
			Select Case i
				Case 0
					ret = rsDiv("WDRules_Anac")
				Case 1
					ret = rsDiv("WDRules_Client")
				Case 2
					ret = rsDiv("WDRules_Delivery")
			End Select
			parse = Split(ret, ";")
			Dim p : p = 0
			For j=0 To 4
				For k=0 To 6
					If p < UBound( parse ) Then
						m_rVal(i,j,k) = CDbl(parse(p))
						p = p + 1
					End If
				Next
			Next
		Next
	End If
	rsDiv.Close()

	' Salva valores
	If cls = "true" Then
		Dim exec(3) : exec(0) = "": exec(1) = "": exec(2) = ""
		Dim save : save = False
		For i=0 To 2
			For j=0 To 4
				For k=0 To 6
					ret = Request.Cookies(tag & CStr(i) & CStr(j) & CStr(k))
					If ret <> "" Then
						' Para limpar: Response.Cookies(tag & CStr(i) & CStr(j) & CStr(k)).Expires = DateAdd("d",-1,Date())
						ret = CDbl(ret)
						If m_rVal(i,j,k) <> ret Then
							m_rVal(i,j,k) = ret
							save = True
						End If
					End If
					exec(i) = exec(i) & CStr(m_rVal(i,j,k)) & ";"
				Next
			Next
		Next
		If save = True Then
			querySQL = "UPDATE WDRules SET WDRules_Anac='" & exec(0) & "', WDRules_Client='" & exec(1) & "', WDRules_Delivery='" & exec(2) & "' WHERE WDRules_Id = '" & tag & "'"
			ret = oDbFDH.Execute( querySQL )
		End If
		oDbFDH.Close() %>
		<script type="text/javascript">
			window.close();		
		</script>
		<%
		Response.End
	End If

	oDbFDH.Close()
 %>
<html>
<head>
<title>Regras para cálculo dos indicadores de atraso</title>
<script type="text/javascript">

	function Init() {
		var x;
		<%
		Dim i, j, k
		For i=0 To 2
			For j=0 To 4
				For k=0 To 6 %>
		x = document.getElementById("R<%=i %><%=j %><%=k%>");
		if (x != null) {
			x.value = "<%=m_rVal(i, j, k) %>";
		}
				<%
				Next
			Next
		Next %>
	}

	function isNumberKey(evt) {
		var charCode = (evt.which) ? evt.which : evt.keyCode;
		if (charCode != 44 && charCode > 31
            && (charCode < 48 || charCode > 57))
			return false;
		return true;
	}

	function setVal(obj) {
		var name = obj.id.substring(1);
		var value = obj.value;
		document.cookie = "<%=tag %>" + name + "=" + value;
	}

	function closeIt() {
		window.open("hWDRulesCfg.asp?tag=<%=tag %>&close=true","_blank","width=10,height=10");
		return "Favor confirmar fechamento da tela de configurações das regras.";
	}
	<%
	If cls <> "true" Then %>
	window.onbeforeunload = closeIt;
	<%
	End If %>

</script>
</head>

<body OnLoad="Init();">

<form method="POST">
    <pre>
    Regras para os cálculos dos indicadores de atraso:

        '--------------------------------------------------------------------------
        ' resDelayANAC	- Indica atraso da ANAC na analise dos processos
        If oAIRData.ANAC = 0 Then
            resDelayANAC = RES_NONE

        ' RES_CRITICO
        ElseIf oAIRData.ANAC > <input type="text" size="5" id="R000" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ANACDelay > <input type="text" size="5" id="R001" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ANACMaxDays > <input type="text" size="5" id="R002" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.ANAC > <input type="text" size="5" id="R003" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC60d/oAIRData.ANAC > <input type="text" size="5" id="R004" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.ANAC > <input type="text" size="5" id="R005" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC30d/oAIRData.ANAC > <input type="text" size="5" id="R006" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayANAC = RES_CRITICO


        ' RES_RUIM
        ElseIf oAIRData.ANAC > <input type="text" size="5" id="R010" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ANACDelay > <input type="text" size="5" id="R011" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ANACMaxDays > <input type="text" size="5" id="R012" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.ANAC > <input type="text" size="5" id="R013" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC60d/oAIRData.ANAC > <input type="text" size="5" id="R014" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.ANAC > <input type="text" size="5" id="R015" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC30d/oAIRData.ANAC > <input type="text" size="5" id="R016" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayANAC = RES_RUIM

        ' RES_ALERTA
        ElseIf oAIRData.ANAC > <input type="text" size="5" id="R020" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ANACDelay > <input type="text" size="5" id="R021" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ANACMaxDays > <input type="text" size="5" id="R022" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.ANAC > <input type="text" size="5" id="R023" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC60d/oAIRData.ANAC > <input type="text" size="5" id="R024" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.ANAC > <input type="text" size="5" id="R025" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC30d/oAIRData.ANAC > <input type="text" size="5" id="R026" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayANAC = RES_ALERTA

        ' RES_BOM
        ElseIf oAIRData.ANAC > <input type="text" size="5" id="R030" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ANACDelay > <input type="text" size="5" id="R031" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ANACMaxDays > <input type="text" size="5" id="R032" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.ANAC > <input type="text" size="5" id="R033" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC60d/oAIRData.ANAC > <input type="text" size="5" id="R034" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.ANAC > <input type="text" size="5" id="R035" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC30d/oAIRData.ANAC > <input type="text" size="5" id="R036" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayANAC = RES_BOM

        ' RES_OTIMO
        ElseIf oAIRData.ANAC > <input type="text" size="5" id="R040" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ANACDelay > <input type="text" size="5" id="R041" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ANACMaxDays > <input type="text" size="5" id="R042" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.ANAC > <input type="text" size="5" id="R043" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC60d/oAIRData.ANAC > <input type="text" size="5" id="R044" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.ANAC > <input type="text" size="5" id="R045" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.ANAC30d/oAIRData.ANAC > <input type="text" size="5" id="R046" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayANAC = RES_OTIMO

        ' RES_NONE
        Else
            resDelayANAC = RES_NONE
        End If


        '--------------------------------------------------------------------------
        ' resDelayClient    - Indica atraso da Empresa/Cliente na resposta dos processos
        If oAIRData.Client = 0 Then
            resDelayClient = RES_NONE

        ' RES_CRITICO
        ElseIf oAIRData.Client > <input type="text" size="5" id="R100" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ClientDelay > <input type="text" size="5" id="R101" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ClientMaxDays > <input type="text" size="5" id="R102" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Then
                resDelayClient = RES_CRITICO

        ' RES_RUIM
        ElseIf oAIRData.Client > <input type="text" size="5" id="R110" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ClientDelay > <input type="text" size="5" id="R111" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ClientMaxDays > <input type="text" size="5" id="R112" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Then
                resDelayClient = RES_RUIM

        ' RES_ALERTA
        ElseIf oAIRData.Client > <input type="text" size="5" id="R120" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ClientDelay > <input type="text" size="5" id="R121" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ClientMaxDays > <input type="text" size="5" id="R122" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Then
                resDelayClient = RES_ALERTA

        ' RES_BOM
        ElseIf oAIRData.Client > <input type="text" size="5" id="R130" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ClientDelay > <input type="text" size="5" id="R131" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ClientMaxDays > <input type="text" size="5" id="R132" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Then
                resDelayClient = RES_BOM

        ' RES_OTIMO
        ElseIf oAIRData.Client > <input type="text" size="5" id="R140" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.ClientDelay > <input type="text" size="5" id="R141" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.ClientMaxDays > <input type="text" size="5" id="R142" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Then
                resDelayClient = RES_OTIMO

        ' RES_NONE
        Else
            resDelayClient = RES_NONE
        End If


        '--------------------------------------------------------------------------
        ' resDelayDelivery    - Indica atraso na distribuição dos processos
        If oAIRData.Delivery = 0 Then
            resDelayDelivery = RES_NONE

        ' RES_CRITICO
        ElseIf oAIRData.Delivery > <input type="text" size="5" id="R200" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.DeliveryDelay > <input type="text" size="5" id="R201" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.DeliveryMaxDays > <input type="text" size="5" id="R202" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.Delivery > <input type="text" size="5" id="R203" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery14d/oAIRData.Delivery > <input type="text" size="5" id="R204" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.Delivery > <input type="text" size="5" id="R205" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery7d/oAIRData.Delivery > <input type="text" size="5" id="R206" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayDelivery = RES_CRITICO

        ' RES_RUIM
        ElseIf oAIRData.Delivery > <input type="text" size="5" id="R210" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.DeliveryDelay > <input type="text" size="5" id="R211" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.DeliveryMaxDays > <input type="text" size="5" id="R212" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.Delivery > <input type="text" size="5" id="R213" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery14d/oAIRData.Delivery > <input type="text" size="5" id="R214" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.Delivery > <input type="text" size="5" id="R215" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery7d/oAIRData.Delivery > <input type="text" size="5" id="R216" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayDelivery = RES_RUIM

        ' RES_ALERTA
        ElseIf oAIRData.Delivery > <input type="text" size="5" id="R220" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.DeliveryDelay > <input type="text" size="5" id="R221" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 oAIRData.DeliveryMaxDays > <input type="text" size="5" id="R222" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.Delivery > <input type="text" size="5" id="R223" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery14d/oAIRData.Delivery > <input type="text" size="5" id="R224" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.Delivery > <input type="text" size="5" id="R225" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery7d/oAIRData.Delivery > <input type="text" size="5" id="R226" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayDelivery = RES_ALERTA

        ' RES_BOM
        ElseIf oAIRData.Delivery > <input type="text" size="5" id="R230" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.DeliveryDelay > <input type="text" size="5" id="R231" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.DeliveryMaxDays > <input type="text" size="5" id="R232" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.Delivery > <input type="text" size="5" id="R233" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery14d/oAIRData.Delivery > <input type="text" size="5" id="R234" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.Delivery > <input type="text" size="5" id="R235" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery7d/oAIRData.Delivery > <input type="text" size="5" id="R236" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayDelivery = RES_BOM

        ' RES_OTIMO
        ElseIf oAIRData.Delivery > <input type="text" size="5" id="R240" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
               ( oAIRData.DeliveryDelay > <input type="text" size="5" id="R241" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _
                 oAIRData.DeliveryMaxDays > <input type="text" size="5" id="R242" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> Or _ 
                 ( oAIRData.Delivery > <input type="text" size="5" id="R243" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery14d/oAIRData.Delivery > <input type="text" size="5" id="R244" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) Or _
                 ( oAIRData.Delivery > <input type="text" size="5" id="R245" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> And _
                   oAIRData.Delivery7d/oAIRData.Delivery > <input type="text" size="5" id="R246" OnKeyPress="return isNumberKey(event);" OnKeyUp="setVal(this);"<%=readonly %>> ) ) Then
                resDelayDelivery = RES_OTIMO

        ' RES_NONE
        Else
            resDelayDelivery = RES_NONE
        End If

	' ATENÇÃO: A edição dos dados deve ser solicitada aos usuários com perfil de Administrador

    </pre>
</form>
</body>
</html>
