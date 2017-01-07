<html>
<head>
<title>get variables</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<%
'Loop through the server variables collection
' with filtering option

'Set whether ALL type server variables are filtered oit
Dim FilterAllType
FilterAllType = True

'Now loop through all the server variables
Dim ServerVar, tmpValue
For Each ServerVar In Request.ServerVariables
    If (InStr(ServerVar,"_ALL") + InStr(ServerVar,"ALL_") = 0) OR _
      Not FilterAllType Then
        tmpValue = Request.ServerVariables(ServerVar)
        Response.Write ServerVar & " = " & tmpValue & "<br />"
    End If
Next

%>

<body>

</body>
</html>
