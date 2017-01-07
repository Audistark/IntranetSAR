<%
Class cWebService

    Public WSDL         
    'http://sdadf1004.anac.gov.br/ProtonWS/proton.asmx?WSDL
    'http://sdadf1004.anac.gov.br/wsSIGEC/sMultas.asmx?WSDL
    'http://sei-lab.anac.gov.br/sei/controlador_ws.php?servico=sei

    Public oSoapClient
    Public tagName

    Public numResult
    Public recordSetLimit
    Public NumElements
    Public Query
    Public NameSpace


	'----------------------------------------------------------------------------------------------------
	' Descrição: Consumir métodos dos web services.
    ' Parâmetros: .
    ' Retorno: .
    ' Retorno erro: .
    Public Function Invocar(NameSpace, Method, values)

        set xmlDom  = Server.CreateObject("MSXML2.DOMDocument")
        set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")

        set Envelope = xmlDom.CreateElement("soap:Envelope")
            Envelope.setAttribute "xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance"
            Envelope.setAttribute "xmlns:xsd","http://www.w3.org/2001/XMLSchema"
            Envelope.setAttribute "xmlns:soap","http://schemas.xmlsoap.org/soap/envelope/"

        xmlDom.documentElement = Envelope

        set Soap = xmlDom.CreateElement("soap:Body")
            Envelope.appendChild(soap)

        set body = xmlDom.CreateElement(Method)
           body.setAttribute "xmlns", NameSpace

        dim NumValues : NumValues = UBound(values, 1)

        dim x
        for x = 0 to NumValues 
            set element                 = xmlDom.createELement(values(x, 0))
                element.dataType        = values(x, 1)
                element.nodeTypedValue  = values(x, 2)
                body.appendChild(element)
        next

        soap.appendChild(body)

        xmlhttp.open "POST", WSDL, false
        xmlhttp.setRequestHeader "SOAPAction", NameSpace & Method
        xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"

        On Error Resume Next
        xmlhttp.send(xmldom.xml)
        if Err.number < 0 then
            Response.Clear()
            Response.Write( "Web Service Error<br>" )
            Response.Write( "Err.number: 0x" & Hex(Err.number) & "<br />" )
            Response.Write( "Err.description: " & Err.Description & "<br /><br />" )
            Response.End
        end if

        On Error Goto 0

        If xmlhttp.status = 200 then

            dim Result, Result2, Query
            ' Result = Resposta XMLHTTp
            ' Result2 = DOMXML
            ' Query = Query XML

            Result = xmlhttp.responseXML.xml
            xmldom.loadXML(Result)

            Query = "soap:Envelope/soap:Body/" & Method & "Response/" & Method & "Result"

            if xmldom.selectNodes(Query & "/xs:schema/xs:element").length >= 1 then

                me.tagName = xmldom.selectNodes(Query & "/xs:schema/xs:element/xs:complexType/xs:choice/xs:element").item(0).attributes.item(0).text

                if xmldom.getElementsByTagName(me.tagName).Length = 0 then
                    me.numResult = 0
                    me.tagName = ""
                    Set Invocar = tratarXML(xmldom.selectNodes(Query).item(0).text, "")
                else
                    me.NumElements = xmldom.selectNodes(Query & "/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/xs:element").length
                    me.numResult = xmldom.getElementsByTagName(me.tagName).Length
                    Set Invocar = tratarXML(xmldom,Query)
                end if

            else 
                me.numResult = 0
                me.tagName = ""
                Set Invocar = tratarXML(xmldom.selectNodes(Query).item(0).text, "")
            end if

        Else

            Response.Clear()
            Response.write "Status: " & xmlhttp.status & "<br />"
            Response.write "Error: " & xmlhttp.responseText & "<br /><br />"
            Response.write xmldom.xml
            Response.End

            me.numResult = -xmlhttp.status

            set Invocar = me

        End If

        set xmlDom = nothing
        set xmlhttp = nothing
        set Envelope = nothing

    End Function

	'----------------------------------------------------------------------------------------------------
	' Descrição: Tratar resultado do xml retornado pelo web service.
    ' Parâmetros: .
    ' Retorno: .
    ' Retorno erro: .
    Private Function tratarXML(Result,Query)

        if me.tagName <> "" then
            Length1 = Result.getElementsByTagName(tagName).Length - 1
            Length2 = Result.getElementsByTagName(tagName).item(0).childNodes.Length - 1
        end if
                
        dim oRecordset

        set oRecordset = Server.CreateObject("ADODB.Recordset")

        set oRecordset.ActiveConnection = nothing
            oRecordset.CursorLocation = 3
            oRecordset.CursorType = 3
            oRecordset.LockType = 4

        dim QueryRecordSet, b
        QueryRecordset = Query & "/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/xs:element"

        dim ColumnsName(50), ColunaRepetida


        for c = 0 To NumElements - 1                    
            ColumnsName(c) = Result.selectNodes(QueryRecordSet).item(c).attributes.item(0).text
        next

        for c = 1 to NumElements -1

            for x = 0 to NumElements - 1

                if x <> c then
                    if ColumnsName(c) = ColumnsName(x) then
                    ColumnsName(c) = ColumnsName(c) & "_"
                    ColunaRepetida = true
                    end if
                end if

            next

        next

        if tagName <> "" then

            with oRecordset.Fields
                .Append "ID", 8
                    for c=0 to NumElements - 1                    
                        .Append ColumnsName(c), 8
                    next                        
            end with

            oRecordset.open

            for I=0 to Length1

                oRecordset.addNew 
                oRecordset.fields("ID") = I+1

                for c=0 to Result.getElementsByTagName(tagName).item(I).childNodes.Length - 1
                    if ColunaRepetida then
                        oRecordset.fields(ColumnsName(c)) = Result.getElementsByTagName(tagName).item(I).childNodes(c).nodeTypedValue  
                    else
                        oRecordset.fields(Result.getElementsByTagName(tagName).item(I).childNodes(c).nodeName) = Result.getElementsByTagName(tagName).item(I).childNodes(c).nodeTypedValue   
                    end if 
                next

                oRecordset.MoveFirst 

            next

        else

            with oRecordset.Fields
                .Append "ID", 8
                .Append "Result",8
            end with

            oRecordset.open 

            oRecordset.addNew 
            oRecordset.fields("ID") = 1
            oRecordset.fields("Result") = Result
            oRecordset.MoveFirst 

        end if

        set tratarXML = oRecordset
        set oRecordset = nothing

    End Function

End Class
%>