<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.CodePage = 1252 %> 
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!--#include virtual = "/inet2/class/cAAA.asp" -->

<%

Private Sub alert(msg)
%>
<html><head>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1"/>
<script language="JavaScript" type="text/javascript">
	alert('<%=msg %>');
</script>
</head><body></body></html>
<%
End Sub

Private Sub goBack()
%>
<script language="JavaScript" type="text/javascript">
	history.back();
</script>
<%
	Response.End
End Sub

Private Sub closeWin()
%>
<script language="JavaScript" type="text/javascript">
	window.close();
</script>
<%
	Response.End
End Sub

'---------------------------------------------------------------------------------------
' decodeBase64(base64)
Private Function decodeBase64(base64)
	Dim DM, EL
	Set DM = CreateObject("Microsoft.XMLDOM")
	' Create temporary node with Base64 data type
	Set EL = DM.createElement("tmp")
	EL.DataType = "bin.base64"
	' Set encoded String, get bytes
	'
	' "data:image/png;base64,iVBORw0KG ... Jgg=="
	Dim b64 : b64 = Mid(base64,23,Len(base64)-22)
	EL.Text = b64
	decodeBase64 = EL.NodeTypedValue
End Function


'----------------------------------------------------------
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
' Windows Authentication
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	oAAA.Print()
End If

' http://sar/inet2/stats/ChartImageUpld.asp?img=

Dim binarydata, strHDLocation, objADOStream, objFSO
Dim name, img64

' File 1
name = Request.Form("filename_1")
img64 = Request.Form("image64_1")

If name <> "" And img64 <> "" Then

	binarydata = decodeBase64(img64)

	' Set your settings
	strHDLocation = Request.ServerVariables( "APPL_PHYSICAL_PATH" ) & _
					"Public\img_" & Year(Date()) & Month(Date()) & Day(Date()) & "_" & name & ".png"

	Set objADOStream = CreateObject("ADODB.Stream")

	objADOStream.Open
	objADOStream.Type = 1 'adTypeBinary

	objADOStream.Write binarydata
	objADOStream.Position = 0 'Set the stream position to the start

	Set objFSO = Createobject("Scripting.FileSystemObject")
	If objFSO.Fileexists(strHDLocation) Then
		objFSO.DeleteFile strHDLocation
	End If
	Set objFSO = Nothing

	objADOStream.SaveToFile strHDLocation
	objADOStream.Close
	Set objADOStream = Nothing

End If

' File 2
name = Request.Form("filename_2")
img64 = Request.Form("image64_2")

If name <> "" And img64 <> "" Then

	binarydata = decodeBase64(img64)

	' Set your settings
	strHDLocation = Request.ServerVariables( "APPL_PHYSICAL_PATH" ) & _
					"Public\img_" & Year(Date()) & Month(Date()) & Day(Date()) & "_" & name & ".png"

	Set objADOStream = CreateObject("ADODB.Stream")

	objADOStream.Open
	objADOStream.Type = 1 'adTypeBinary

	objADOStream.Write binarydata
	objADOStream.Position = 0 'Set the stream position to the start

	Set objFSO = Createobject("Scripting.FileSystemObject")
	If objFSO.Fileexists(strHDLocation) Then
		objFSO.DeleteFile strHDLocation
	End If
	Set objFSO = Nothing

	objADOStream.SaveToFile strHDLocation
	objADOStream.Close
	Set objADOStream = Nothing

End If

closeWin()

%>








