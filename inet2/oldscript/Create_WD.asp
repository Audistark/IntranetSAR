<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<%
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	oAAA.Print()
End If

' Só MASTER
If oAAA.AuthorWinMaster() <> True Then
	Response.Status = "403 Forbidden"
	Response.End
End If

Dim querySQL, rsDiv
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Response.Status = "200 OK"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set ret = oDbFDH.Execute( "DELETE * FROM MPH WHERE MPH_NUM='000';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM MPH_Cancela WHERE MPH_NUM='000';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM MPH_Rev WHERE MPH_NUM='000';" )

'Set ret = oDbFDH.Execute( "DROP TABLE AIR145Statistics;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDInclEO145;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDManuals145;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDCert145;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDOthers145;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDInclEO135;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDManuals135;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDOthers135;" )
'Set ret = oDbFDH.Execute( "DROP TABLE WDControl;" )

'''''''''''''''''''' Create Table AIRStatistics ''''''''''''''''''''''''''''''''''
'
'On Error Resume Next
'Set ret = oDbFDH.Execute( "CREATE TABLE AIRStatistics " & _
'						  "	(AIRStats_DATE DATE NOT NULL, " & _
'						  "	 AIRStats_SOLIC TEXT(3) NOT NULL, " & _
'						  "	 AIRStats_GTAR TEXT(10) NOT NULL, " & _
'						  "	 AIRStats_RBAC TEXT(10) NOT NULL, " & _
'						  "	 AIRStats_ANAC INTEGER, " & _
'						  "	 AIRStats_30D_ANAC INTEGER, " & _
'						  "	 AIRStats_60D_ANAC INTEGER, " & _
'						  "	 AIRStats_MAX_ANAC INTEGER, " & _
'						  "	 AIRStats_30D_CLOSED INTEGER, " & _
'						  "	 AIRStats_30D_DOCS INTEGER, " & _
'						  "	 AIRStats_CLIENT INTEGER, " & _
'						  "	 AIRStats_DELAYED_CLIENT INTEGER, " & _
'						  "	 AIRStats_MAX_CLIENT INTEGER, " & _
'						  "	 AIRStats_DELIVERY INTEGER, " & _
'						  "	 AIRStats_7D_DELIVERY INTEGER, " & _
'						  "	 AIRStats_14D_DELIVERY INTEGER, " & _
'						  "	 AIRStats_MAX_DELIVERY INTEGER, " & _
'						  "  AIRStats_TIMESTAMP DATETIME, " & _
'						  "	 CONSTRAINT AIRStatistics_PK PRIMARY KEY(AIRStats_DATE,AIRStats_GTAR,AIRStats_RBAC,AIRStats_SOLIC));" )
'If ret Is Nothing Then
'	Response.Write "The TABLE AIRStatistics already exists in Database.<br>"
'Else
'	Response.Write "The TABLE AIRStatistics was created sucessfully in Database!<br>"
'End If
'On Error GoTo 0
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''' Create Table WDRules''''''''''''''''''''''''''''''''''
'
'On Error Resume Next
'Set ret = oDbFDH.Execute( "CREATE TABLE WDRules " & _
'						  "	(WDRules_Id TEXT(20) NOT NULL, " & _
'						  "	 WDRules_Anac TEXT NOT NULL, " & _
'						  "	 WDRules_Client TEXT NOT NULL, " & _
'						  "	 WDRules_Delivery TEXT NOT NULL, " & _
'						  "	 CONSTRAINT AIRStatistics_PK PRIMARY KEY(WDRules_Id));" )
'If ret Is Nothing Then
'	Response.Write "The TABLE WDRules already exists in Database.<br>"
'Else
'	Response.Write "The TABLE WDRules was created sucessfully in Database!<br>"
'End If
'On Error GoTo 0

'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145002';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145005';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145006';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145008';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145009';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Manuals145';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='InclEO145';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145003';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145007';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145013';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145014';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='RCA91';" )


'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others135';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='InclEO135';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Manuals135';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others135006';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others135007';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others135008';" )
'Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others135023';" )

'------------------------------------------------------------------
'
'						  "	(WDRules_Id TEXT(20) NOT NULL, " & _
'						  "	 WDRules_Anac TEXT NOT NULL, " & _
'						  "	 WDRules_Client TEXT NOT NULL, " & _
'						  "	 WDRules_Delivery TEXT NOT NULL, " & _
'						  "	 CONSTRAINT AIRStatistics_PK PRIMARY KEY(WDRules_Id));" )
'If ret Is Nothing Then
'	Response.Write "The TABLE WDRules already exists in Database.<br>"

querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145002DF', '0;3;60;0;0;0;1;0;1;45;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;3;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145005DF', '0;8;60;0;0;0;1;0;4;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;5;14;0;0;0;1;0;3;10;0;0;0;1;0;1;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145006DF', '0;3;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145008DF', '0;3;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145009DF', '0;2;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'InclEO135DF', '0;3;45;0;1;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;2;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'InclEO145DF', '0;3;45;0;0;0;1;0;2;30;0;0;0;1;0;1;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;180;-1;-1;-1;-1;0;3;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;4;14;0;0;0;1;0;2;10;0;0;0;1;0;1;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Manuals135DF', '0;10;90;0;1;0;1;0;5;60;0;1;0;1;0;0;30;0;0;0;1;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;10;45;0;1;0;1;0;6;30;0;1;0;1;0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Manuals145DF', '0;3;90;0;1;0;1;0;2;60;0;0;0;1;0;1;30;0;0;0;0;0;0;20;0;0;0;0;0;0;15;0;0;0;0;', '0;5;150;-1;-1;-1;-1;0;4;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;2;2;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;9;45;0;1;0;1;0;6;30;0;1;0;1;0;4;15;0;1;0;1;0;2;10;0;0;0;1;0;0;7;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'MEL135DF', '0;3;45;0;1;0;1;0;0;30;0;0;0;1;0;0;15;0;0;0;0;0;0;8;0;0;0;0;0;0;4;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135007DF', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135008DF', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135023DF', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145003DF', '0;4;90;0;1;0;1;0;2;60;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145007DF', '0;10;90;0;1;0;1;0;4;45;0;0;0;1;0;2;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;10;120;-1;-1;-1;-1;0;4;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145013DF', '0;5;90;0;1;0;1;0;2;60;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;4;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145014DF', '0;4;60;0;1;0;1;0;2;45;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'PM135DF', '0;3;60;0;1;0;1;0;2;45;0;0;0;1;0;0;30;0;0;0;1;0;0;14;0;0;0;0;0;0;4;0;0;0;0;', '0;10;360;-1;-1;-1;-1;0;5;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'RCA91DF', '0;0;9;0;0;0;0;0;0;7;0;0;0;0;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;2;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;0;7;0;0;0;0;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;2;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )


querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145002RJ', '0;3;60;0;0;0;1;0;1;45;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;3;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145005RJ', '0;8;60;0;0;0;1;0;4;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;5;14;0;0;0;1;0;3;10;0;0;0;1;0;1;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145006RJ', '0;3;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145008RJ', '0;3;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145009RJ', '0;2;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'InclEO135RJ', '0;3;45;0;1;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;2;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'InclEO145RJ', '0;3;45;0;0;0;1;0;2;30;0;0;0;1;0;1;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;180;-1;-1;-1;-1;0;3;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;4;14;0;0;0;1;0;2;10;0;0;0;1;0;1;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Manuals135RJ', '0;10;90;0;1;0;1;0;5;60;0;1;0;1;0;0;30;0;0;0;1;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;10;45;0;1;0;1;0;6;30;0;1;0;1;0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Manuals145RJ', '0;3;90;0;1;0;1;0;2;60;0;0;0;1;0;1;30;0;0;0;0;0;0;20;0;0;0;0;0;0;15;0;0;0;0;', '0;5;150;-1;-1;-1;-1;0;4;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;2;2;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;9;45;0;1;0;1;0;6;30;0;1;0;1;0;4;15;0;1;0;1;0;2;10;0;0;0;1;0;0;7;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'MEL135RJ', '0;3;45;0;1;0;1;0;0;30;0;0;0;1;0;0;15;0;0;0;0;0;0;8;0;0;0;0;0;0;4;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135007RJ', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135008RJ', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135023RJ', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )

oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145003RJ', '0;4;90;0;1;0;1;0;2;60;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145007RJ', '0;10;90;0;1;0;1;0;4;45;0;0;0;1;0;2;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;10;120;-1;-1;-1;-1;0;4;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145013RJ', '0;5;90;0;1;0;1;0;2;60;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;4;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145014RJ', '0;4;60;0;1;0;1;0;2;45;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'PM135RJ', '0;3;60;0;1;0;1;0;2;45;0;0;0;1;0;0;30;0;0;0;1;0;0;14;0;0;0;0;0;0;4;0;0;0;0;', '0;10;360;-1;-1;-1;-1;0;5;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'RCA91RJ', '0;0;9;0;0;0;0;0;0;7;0;0;0;0;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;2;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;0;7;0;0;0;0;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;2;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )


querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145002SP', '0;3;60;0;0;0;1;0;1;45;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;3;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145005SP', '0;8;60;0;0;0;1;0;4;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;5;14;0;0;0;1;0;3;10;0;0;0;1;0;1;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145006SP', '0;3;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145008SP', '0;3;45;0;0;0;1;0;2;30;0;0;0;0;0;1;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;2;14;0;0;0;1;0;1;10;0;0;0;1;0;0;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Cert145009SP', '0;2;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;15;0;0;0;0;0;0;7;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;1;0;0;4;0;0;0;0;0;0;2;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'InclEO135SP', '0;3;45;0;1;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;2;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'InclEO145SP', '0;3;45;0;0;0;1;0;2;30;0;0;0;1;0;1;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;180;-1;-1;-1;-1;0;3;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;4;14;0;0;0;1;0;2;10;0;0;0;1;0;1;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Manuals135SP', '0;10;90;0;1;0;1;0;5;60;0;1;0;1;0;0;30;0;0;0;1;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;10;45;0;1;0;1;0;6;30;0;1;0;1;0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Manuals145SP', '0;3;90;0;1;0;1;0;2;60;0;0;0;1;0;1;30;0;0;0;0;0;0;20;0;0;0;0;0;0;15;0;0;0;0;', '0;5;150;-1;-1;-1;-1;0;4;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;2;2;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;9;45;0;1;0;1;0;6;30;0;1;0;1;0;4;15;0;1;0;1;0;2;10;0;0;0;1;0;0;7;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'MEL135SP', '0;3;45;0;1;0;1;0;0;30;0;0;0;1;0;0;15;0;0;0;0;0;0;8;0;0;0;0;0;0;4;0;0;0;0;', '0;20;360;-1;-1;-1;-1;0;10;180;-1;-1;-1;-1;0;5;90;-1;-1;-1;-1;0;3;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135007SP', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135008SP', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others135023SP', '0;3;45;0;0;0;1;0;0;30;0;0;0;1;0;0;20;0;0;0;0;0;0;10;0;0;0;0;0;0;5;0;0;0;0;', '0;5;360;-1;-1;-1;-1;0;4;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;', '0;2;14;0;1;0;1;0;1;9;0;0;0;1;0;0;7;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145003SP', '0;4;90;0;1;0;1;0;2;60;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145007SP', '0;10;90;0;1;0;1;0;4;45;0;0;0;1;0;2;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;10;120;-1;-1;-1;-1;0;4;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145013SP', '0;5;90;0;1;0;1;0;2;60;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;4;120;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'Others145014SP', '0;4;60;0;1;0;1;0;2;45;0;0;0;1;0;0;30;0;0;0;0;0;0;15;0;0;0;0;0;0;5;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;1;30;-1;-1;-1;-1;0;0;15;-1;-1;-1;-1;', '0;3;14;0;0;0;1;0;2;10;0;0;0;1;0;1;7;0;0;0;0;0;0;4;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'PM135SP', '0;3;60;0;1;0;1;0;2;45;0;0;0;1;0;0;30;0;0;0;1;0;0;14;0;0;0;0;0;0;4;0;0;0;0;', '0;10;360;-1;-1;-1;-1;0;5;180;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;3;14;0;1;0;1;0;1;9;0;0;0;1;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO WDRules (WDRules_Id, WDRules_Anac, WDRules_Client, WDRules_Delivery)" & _
			"    VALUES ( 'RCA91SP', '0;0;9;0;0;0;0;0;0;7;0;0;0;0;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;2;0;0;0;0;', '0;5;120;-1;-1;-1;-1;0;3;90;-1;-1;-1;-1;0;2;90;-1;-1;-1;-1;0;1;60;-1;-1;-1;-1;0;0;45;-1;-1;-1;-1;', '0;0;7;0;0;0;0;0;0;5;0;0;0;0;0;0;3;0;0;0;0;0;0;2;0;0;0;0;0;0;1;0;0;0;0;' )"
oDbFDH.Execute( querySQL )

Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145002';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145005';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145006';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145008';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Cert145009';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='InclEO135';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='InclEO145';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Manuals135';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Manuals145';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='MEL135';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others135';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145003';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145007';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145013';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='Others145014';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='PM135';" )
Set ret = oDbFDH.Execute( "DELETE * FROM WDRules WHERE WDRules_Id='RCA91';" )

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.Status = "200 OK"
Response.Write "Data was inserted sucessfully in Database!"
Response.End

%>