<%
Class cWebService

    Public WSDL         
    'http://sdadf1004.anac.gov.br/ProtonWS/proton.asmx?WSDL
    'http://sdadf1004.anac.gov.br/wsSIGEC/sMultas.asmx?WSDL
    'http://sei-lab.anac.gov.br/sei/controlador_ws.php?servico=sei

    Public numResult

    '----------------------------------------------------------------------------------------------------
    ' Descrição: Consumir métodos dos web services.
    ' Parâmetros: .
    ' Retorno: .
    ' Retorno erro: .
    Public Function Invocar(NameSpace, Method, values, results)

        Set xmlDoc  = Server.CreateObject("MSXML2.DOMDocument.6.0")
        Set xmlHttp = server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
        Set Envelope = xmlDoc.CreateElement("soap:Envelope")
            Envelope.setAttribute "xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance"
            Envelope.setAttribute "xmlns:xsd","http://www.w3.org/2001/XMLSchema"
            Envelope.setAttribute "xmlns:soap","http://schemas.xmlsoap.org/soap/envelope/"

        xmlDoc.documentElement = Envelope

        Set Soap = xmlDoc.CreateElement("soap:Body")
		Envelope.appendChild(soap)

        Set body = xmlDoc.CreateElement(Method)
		body.setAttribute "xmlns", NameSpace

        Dim NumValues : NumValues = UBound(values, 1)

		Dim oRecordSet
		Set oRecordSet = Server.CreateObject("ADODB.Recordset")
		Set oRecordSet.ActiveConnection = nothing
			oRecordSet.CursorLocation = 3
			oRecordSet.CursorType = 3
			oRecordSet.LockType = 4

		Dim i
		For i=0 To UBound(results)
			oRecordSet.Fields.Append results(i), 8
		Next

        Dim x
'		Response.Write( "Params: Value(" & NumValues & "," & UBound(values, 2) & ")<br>" )
        For x = 0 to NumValues 
            Set element                 = xmlDoc.createELement(values(x, 0))
                element.dataType        = values(x, 1)
                element.nodeTypedValue  = values(x, 2)
                body.appendChild(element)
'                Response.Write( "Value(" & x & ",0) = " & values(x, 0) & "<br>" )
'                Response.Write( "Value(" & x & ",1) = " & values(x, 1) & "<br>" )
'                Response.Write( "Value(" & x & ",2) = '" & values(x, 2) & "'<br>" )
        Next

        Soap.appendChild(body)

        xmlHttp.open "POST", WSDL, false
        xmlHttp.setRequestHeader "SOAPAction", NameSpace & Method
        xmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"

        On Error Resume Next
        xmlHttp.send(xmlDoc.xml)
        If Err.number < 0 Then
            Response.Clear()
            Response.Write( "Web Service Error<br>" )
            Response.Write( "Err.number: 0x" & Hex(Err.number) & "<br />" )
            Response.Write( "Err.description: " & Err.Description & "<br /><br />" )
            Response.End
        End If
        On Error Goto 0

'	Response.Write( "Http Status = " & xmlHttp.status & "<br>" )

        If xmlHttp.status = 200 Then

			Dim Res
			Res = xmlHttp.responseText

'			Response.Write( "xmlHttp.responseText = " & xmlHttp.responseText & "<br>" )

			xmlDoc.loadXML(Res)
			If xmlDoc.parseError.errorCode <> 0 Then
				Response.Write( "Parse Error: " & xmlDoc.parseError & " - " & myErr.reason & "<br>" )
				Response.End
			End If

			oRecordSet.open
			oRecordSet.addNew

			Dim n : n = 0
			Dim items(20)
			For i=0 To UBound(results)
				Dim str : str = results(i)
				Dim ini : ini = 1
				Dim pos
				Dim j : j = 0
'				Response.Write( "results(" & i & ")='" & str & "'<br>" )
				Do
					pos = InStr(ini,results(i),"/")
					If pos > 0 Then
						str = Mid(results(i),ini,pos-1)
						ini = pos+1
					ElseIf j > 0 Then
						str = Mid(results(i),ini,Len(results(i))-ini+1)
					End If

					On Error Resume Next

					'## Get the items matching the tag:
					If j > 0 Then
						Set items(j) = items(j-1).item(0).getElementsByTagName(str)
					Else
						Set items(j) = xmlDoc.getElementsByTagName(str)
					End If
					If err.number <> 0 Then
						Response.Write( "Fetch error for string '" & str & "' <br>" )
						Response.End
					ElseIf pos = 0 Then
'						Response.Write( "results(" & i & ") = " & items(j).item(0).Text & "<br>" )
						oRecordSet.fields(results(i)) = items(j).item(0).Text
						n = n + 1
					Else
						j = j + 1
					End If

					On Error Goto 0

				Loop While pos > 0

			Next

			oRecordSet.MoveFirst 

'Response.Write( "n = " & n & "<br>" )

			me.numResult = n

			Set Invocar = oRecordSet

        Else

            Response.Clear()
            Response.write "Status: " & xmlHttp.status & "<br />"
            Response.write "Error: " & xmlHttp.responseText & "<br /><br />"
            Response.write xmlDoc.xml
            Response.End

            me.numResult = -xmlHttp.status

            Set Invocar = me

        End If

		Set oRecordSet = Nothing
		Set xmlDoc = Nothing
		Set xmlHttp = Nothing
		Set Envelope = Nothing

    End Function

End Class
%>