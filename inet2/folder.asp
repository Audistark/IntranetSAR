<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	'-------------------------------------------------------------------
	' Usage:
	'
	'	<iframe width='480' height='200' frameborder='0' scrolling='auto' marginwidth='4' marginheight='2'
	'	 id="iFrameDir" src='inet2/folder.asp?Dir=AvGeral\AIR145\Atas&Msg=Atas das Reuniões do AIR145&Width=480'></iframe>
	'

	' declare variables
	Dim objFSO, objFolder
	Dim objCollection, objItem
	Dim strPhysicalPath, strTitle, strServerName
	Dim strLink, frmSize
	Dim strName, strFile, strExt, strAttr
	Dim intSizeB, intSizeK, intAttr, dtmDate

	' declare constants
	Const vbReadOnly = 1
	Const vbHidden = 2
	Const vbSystem = 4
	Const vbVolume = 8
	Const vbDirectory = 16
	Const vbArchive = 32
	Const vbAlias = 64
	Const vbCompressed = 128

	' Get Directory desired
	Dim Dir : Dir = Request.QueryString("Dir")
	Dim Msg : Msg = Request.QueryString("Msg")
	If Dir = "" Or Msg = "" Then
		Response.Status = "400 Bad Request"
		Response.End
	End If
	Dim Width : Width = Request.QueryString("Width")
	If Width < 420 Then
		Width = 420
	End If

	' don't cache the page
	Response.AddHeader "Pragma", "No-Cache"
	Response.CacheControl = "Private"

	' Format final physical path
	strPhysicalPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "FDH\" & Dir

	' build the page title
	strServerName = UCase(Request.ServerVariables("SERVER_NAME"))
	strTitle = "Contents of the " & Dir & " folder"

	' create the file system objects
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(strPhysicalPath)
%>
<!DOCTYPE html>
<html>
<head>
<title><%=strServerName%> - <%=strTitle%></title>
<style>
	body { background: #000000; }
	td
	{
		font-family: Courier New;
		font-size: 12px;
		color: #00FF00;
	}
	a:link {color: #00FF00; text-decoration: none;}
	a:visited {color: #00FF00; text-decoration: none;}
	a:hover {color: #00FF00; text-decoration: none;}
	a:active {color: #00FF00; text-decoration: none;}
</style>
</head>
<body>
<table border="0" cellspacing="0" cellpadding="0">
    <tr>
        <td colspan="9">C:\\<%=Server.HTMLEncode(Msg) %>\</td>
    </tr>
    <tr>
        <td colspan="9">&nbsp;</td>
    </tr>
    <tr>
        <td>File</td>
        <td>&nbsp;</td>
        <td>Date</td>
        <td>&nbsp;</td>
        <td>Time</td>
        <td>&nbsp;</td>
        <td>Att</td>
        <td>&nbsp;</td>
        <td>Size</td>
    </tr>
<%
	Dim lim : lim = CInt(Width/12+0.5)
	Dim lin : lin = String(lim+30,"-")
%>
    <tr>
        <td colspan="9"><%=lin %></td>
    </tr>
<%
	''''''''''''''''''''''''''''''''''''''''
	' output the file list
	''''''''''''''''''''''''''''''''''''''''
	Set objCollection = objFolder.Files
	For Each objItem in objCollection
		strName = objItem.Name
		strLink = "\FDH\" & Dir & "\" & Server.HTMLEncode(Lcase(strName))
		If Len(strName) > lim Then
			strName = Left(strName,lim-2) & ".."
		End If
		intSizeB = objItem.Size
		intSizeK = Int((intSizeB/1024) + .5)
		If intSizeK = 0 Then intSizeK = 1
		If intSizeB < 1000 Then
			frmSize = FormatNumber(intSizeB,0)
		Else
			frmSize = FormatNumber(intSizeK,0) & "K"
		End If
		strAttr = MakeAttr(objItem.Attributes)
		dtmDate = CDate(objItem.DateLastModified)
%>
    <tr>
        <td nowrap="nowrap"><a href="<%=strLink %>" target="_blank"><%=strName %></a></td>
        <td>&nbsp;</td>
        <td><%=FormatDateTime(dtmDate,vbShortDate) %></td>
        <td>&nbsp;</td>
        <td><%=FormatDateTime(dtmDate,vbLongTime) %></td>
        <td>&nbsp;</td>
        <td nowrap="nowrap"><%=strAttr %></td>
        <td>&nbsp;</td>
        <td nowrap="nowrap"><%=frmSize %></td>
    </tr>
<% Next %>
    <tr>
        <td colspan="9">&nbsp;</td>
    </tr>
	<tr>
        <td colspan="9">--End Of Files--</td>
	</tr>
</table>

</body>
</html>
<%
   Set objFSO = Nothing
   Set objFolder = Nothing

   ' this adds the IIf() function to VBScript
   Function IIf(i,j,k)
      If i Then IIf = j Else IIf = k
   End Function

   ' this function creates a string from the file atttributes
   Function MakeAttr(intAttr)
      MakeAttr = MakeAttr & IIf(intAttr And vbArchive,"A","-")
      MakeAttr = MakeAttr & IIf(intAttr And vbSystem,"S","-")
      MakeAttr = MakeAttr & IIf(intAttr And vbHidden,"H","-")
      MakeAttr = MakeAttr & IIf(intAttr And vbReadOnly,"R","-")
   End Function
%>
