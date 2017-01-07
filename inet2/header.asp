<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage = 1252 %> 
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
Dim user : user = oAAA.AuthentWinUser
user = LCase(user)

Dim domain : domain = oAAA.AuthentWinDomain
domain = oAAA.AuthentInetDomain(domain)

<!-- #include virtual = "/inet2/header.html" -->

