<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Option Explicit
Response.Expires = 0

Dim updated : updated = "[Atualizando]"
Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
Dim filename : filename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "FDH\AvGeral\AIR145\ti-update.txt"
On Error Resume Next
Dim fle : Set fle = fso.OpenTextFile(filename)
If Err.Number = 0 Then
	'Do Until fle.AtEndOfStream
	' Last Updated Date: 25/10/2013 13:49:23
	updated = Right( fle.ReadLine, 19)
	'Loop
	fle.Close()
End If
On Error GoTo 0
Set fle = Nothing
Set fso = Nothing
 
Dim title : title = "Demandas Gestão TI-SAR"

%>
<!-- #include file="gti-demands.html" -->

