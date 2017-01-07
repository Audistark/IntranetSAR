<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CodePage = 1252 %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/lib/libFuncDiv.asp" -->
<%

Const version = "1.0/20131123"

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

Private Sub redir(href)
%>
<script language="JavaScript" type="text/javascript">
	<% If href <> "" Then %>
    location.href = "<%=href %>";
	<% End If %>
</script>
<%
	Response.End
End Sub

' declare constants
Const vbReadOnly = 1
Const vbHidden = 2
Const vbSystem = 4
Const vbVolume = 8
Const vbDirectory = 16
Const vbArchive = 32
Const vbAlias = 64
Const vbCompressed = 128

' this adds the IIf() function to VBScript
Private Function IIf(i,j,k)
    If i Then IIf = j Else IIf = k
End Function

' this function creates a string from the file atttributes
Private Function MakeAttr(intAttr)
    MakeAttr = MakeAttr & IIf(intAttr And vbArchive,"A","-")
    MakeAttr = MakeAttr & IIf(intAttr And vbSystem,"S","-")
    MakeAttr = MakeAttr & IIf(intAttr And vbHidden,"H","-")
    MakeAttr = MakeAttr & IIf(intAttr And vbReadOnly,"R","-")
End Function

' show
Private Sub ShowPath(path)

	Dim root : root = Request.ServerVariables("APPL_PHYSICAL_PATH")
	Dim ln : ln = Len(root) + 1

	Dim fs : Set fs = CreateObject("Scripting.FileSystemObject")
	Dim longpath : longpath = root & path
	Dim fl : Set fl = fs.GetFolder(longpath)

	' Diretório atual
	Const white = "white"
	Const pink = "rgb(255, 250, 253)"
	Dim bkgnd : bkgnd = white

	' \\IntranetSAR\src\biometrics\adapters\engine_adapter\objchk_win7_x86\i386
	Dim shwstr : shwstr = folder
	If Len(shwstr) > 42 Then
		shwstr = ".." & Right(shwstr,40)
	End If  

%>
      <tr
 style="text-align: left; height: 25px;">
        <td colspan="2" rowspan="1"
 style="vertical-align: middle; font-family: Courier New;">Directory of \\IntranetSAR\<%=shwstr %></td>
        <td style="font-family: Courier New; width: 200px;">&nbsp;</td>
        <td style="text-align: center; width: 52px;"><img
 type="file" id="newDir" name="newDir"
 src="img/icons/glyphicons_149_folder_new.png"
 onclick="NewFolder(this);"></td>
        <td style="text-align: center; width: 38px;"><label
 for="upldFile" id="button">Upd</label><input type="file"
 id="upldFile" name="upldFile" onchange="uplFile();"></td>
      </tr>
      <tr
 style="text-align: left; height: 25px; background-color: <%=bkgnd %>;">
        <td
 style="width: 42px; font-family: Courier New; text-align: left;">. </td>
        <td style="width: 680px;">&nbsp; </td>
        <td style="font-family: Courier New; width: 200px;">&nbsp;</td>
        <td style="font-family: Courier New; width: 52px;">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
<%

	If path <> "" Then

		' Diretório anterior
		bkgnd = pink
%>
      <tr
 style="text-align: left; height: 25px; background-color: <%=bkgnd %>;">
<%	If path = "" Then %>
        <td
 style="width: 42px; font-family: Courier New; text-align: left;">..</td>
<%	Else
		Dim back : back = ""
		Dim pos : pos = InStrRev(path,"\")
		If pos > 0 Then back = Left(path,pos-1)
		back = Replace(back,"\","/")
 %>
        <td
 style="width: 42px; font-family: Courier New; text-align: left;"><a
 href="JavaScript: goBack('<%=back %>');">..</a></td>
<%	End If %>
        <td style="width: 680px;">&nbsp;</td>
        <td style="font-family: Courier New; width: 200px;">&nbsp;</td>
        <td style="font-family: Courier New; width: 52px;">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
<%
	End If

	' mostra os diretórios
	Dim items, shortdir, nmdir
	Dim dtmDate
	Dim frmSize
	Dim strAttr
	For Each items In fl.SubFolders

		If bkgnd = pink Then
			bkgnd = white
		Else 
			bkgnd = pink
		End If

		shortdir = Mid(items.Path,ln)
		nmdir = Mid(shortdir,Len(path)+1)
		If Left(nmdir,1) = "\" Then nmdir = Mid(nmdir,2)
		shwstr = nmdir
		If Len(shwstr) > 66 Then
			shwstr = Left(shwstr,64) & ".."
		End If  

		'  last modified
		dtmDate = CDate(items.DateLastModified)

%>
      <tr
 style="text-align: left; height: 25px; background-color: <%=bkgnd %>;">
        <td style="width: 42px; text-align: left;"><img
 id="<%=shortdir %>" src="img/icons/glyphicons_144_folder_open.png"
 onclick="setFolder(this);"></td>
        <td style="font-family: Courier New; width: 680px;"><%=shwstr %></td>
        <td style="font-family: Courier New; width: 200px;"><%=FormatDateTime(dtmDate) %></td>
        <td style="font-family: Courier New; width: 52px;">&nbsp;</td>
<%			If items.Files.Count = 0 And items.SubFolders.Count = 0 Then %>
        <td style="text-align: center;  width: 38px;"><img
 style="width: 18px; height: 18px;" alt="del" name="<%=nmdir %>"
 onclick="DelFolder(this);" src="img/icons/glyphicons_207_remove_2.png"></td>
<%			Else %>
        <td>&nbsp;</td>
<%			End If %>
      </tr>
<%
	Next

	' mostra os arquivos
	If fl.Files.Count > 0 Then

		Dim arq
		For Each arq In fl.Files

			If bkgnd = pink Then
				bkgnd = white
			Else 
				bkgnd = pink
			End If

			shortdir = Mid(fl.Path,ln)
			If shortdir <> "" Then shortdir = shortdir & "/"

			shwstr = arq.Name
			If Len(shwstr) > 66 Then
				shwstr = Left(shwstr,64) & ".."
			End If  

			'  last modified
			dtmDate = CDate(arq.DateLastModified)

			' attributes
			strAttr = MakeAttr(arq.Attributes)

%>
      <tr
 style="text-align: left; height: 25px; background-color: <%=bkgnd %>;">
        <td style="width: 42px; text-align: left;"><img
 src="img/icons/glyphicons_036_file.png"></td>
        <td style="font-family: Courier New; width: 680px;"><a
 href="/<%=shortdir %><%=arq.Name %>" target="_blank"><%=shwstr %></a></td>
        <td style="font-family: Courier New; width: 200px;"><%=FormatDateTime(dtmDate) %></td>
        <td style="font-family: Courier New; width: 52px;"><%=strAttr %>&nbsp;</td>
        <td style="text-align: center;  width: 38px;"><img
 style="width: 18px; height: 18px;" alt="del" name="<%=arq.Name %>"
 onclick="Delete(this);" src="img/icons/glyphicons_207_remove_2.png"></td>
      </tr>
<%
		Next
	
	End If

End Sub


	'-------------------------------------------------------------
	'
	'							M  A  I  N
	'

	'-------------------------------------------------------------
	' Init AAA - Authentication, Authorization and Accounting
	Dim oAAA : Set oAAA = new cAAA
	Dim ret : ret = oAAA.WinAuthenticate(True)
	If ret < 0 Then
		Response.Status = "403 Forbidden"
		Response.End
	End If

	' Só MASTER
	If oAAA.AuthorWinMaster() <> True Then
		Response.Status = "403 Forbidden"
		Response.End
	End If


	' Parameters
	Dim oper : oper = ""
	Dim folder : folder = ""
	Dim file : file = ""
	Dim uplFile : uplFile = ""
	Dim filename : filename = ""

	' Smart Upload
	Dim oRequest
	Set oRequest = Server.CreateObject( "AspSmartUpLoad.SmartUpLoad" )
	On Error Resume Next
	oRequest.UpLoad
	If Err.Number = 0 Then
		folder = oRequest.Form("folder")
		folder = Replace(folder,"/","\")
		file = oRequest.Form("file")
		oper = oRequest.Form("oper")
		If Not oRequest.Files.Item(1).IsMissing Then
			uplFile  = oRequest.Files.Item(1).FileName
		End If
	End If
	On Error GoTo 0


	' don't cache the page
	Response.AddHeader "Pragma", "No-Cache"
	Response.CacheControl = "Private"

	Dim objFSO : Set objFSO = Nothing

	Select Case oper

		' Delete
		Case "Del"

			If file <> "" Then

				Set objFSO = Server.CreateObject( "Scripting.FileSystemObject" )
				filename = Request.ServerVariables("APPL_PHYSICAL_PATH") & folder & "\" & file

				On Error Resume Next
				objFSO.DeleteFile filename, True
				If Err.Number <> 0 Then
					alert(Err.Description)
				'Else
				'	alert("operation was executed with success!!!")
				End If
				On Error GoTo 0

				Set objFSO = Nothing

			End If

		' Delete folder
		Case "Dlf"

			If file <> "" Then

				Set objFSO = Server.CreateObject( "Scripting.FileSystemObject" )
				filename = Request.ServerVariables("APPL_PHYSICAL_PATH") & folder & "\" & file

				On Error Resume Next
				objFSO.DeleteFolder filename, True
				If Err.Number <> 0 Then
					alert(Err.Description)
				'Else
				'	alert("operation was executed with success!!!")
				End If
				On Error GoTo 0

				Set objFSO = Nothing

			End If

		' New folder
		Case "New"

			If file <> "" Then

				Set objFSO = Server.CreateObject( "Scripting.FileSystemObject" )
				filename = Request.ServerVariables("APPL_PHYSICAL_PATH") & folder & "\" & file

				On Error Resume Next
				objFSO.CreateFolder filename
				If Err.Number <> 0 Then
					alert(Err.Description)
				'Else
				'	alert("operation was executed with success!!!")
				End If
				On Error GoTo 0

				Set objFSO = Nothing

			End If

		' Upload
		Case "Upl"

			If uplFile <> "" Then

				On Error Resume Next
				oRequest.Save Request.ServerVariables( "APPL_PHYSICAL_PATH" ) & folder
				If Err.Number <> 0 Then
					alert(Err.Description)
				'Else
				'	alert("operation was executed with success!!!")
				End If
				On Error GoTo 0

			End If

		' Show
		'Case Else

	End Select

	Dim title : title = "Intranet SAR Explorer"

 %>
<!DOCTYPE html>
<html>
<head>
  <meta content="text/html; charset=ISO-8859-1"
 http-equiv="content-type">
  <title><%=title %></title>
  <script language="javascript" type="text/javascript">

  	function setFolder(img) {
  		document.getElementById('folder').value = img.id;
  		document.getElementById('oper').value = "Shw";
  		document.root.submit();
  	}

  	function uplFile() {
  		document.getElementById('oper').value = "Upl";
		document.root.submit();
  	}

  	function goBack(path) {
  		document.getElementById('folder').value = path;
  		document.getElementById('oper').value = "Shw";
  		document.root.submit();
  	}

  	function Delete(file) {
  		var msg = "Favor confirmar que o arquivo a seguir será deletado:\n" + document.getElementById('folder').value + "\\" + file.name;
  		var res = confirm(msg);
  		if (res == true) {
  			document.getElementById('file').value = file.name;
  			document.getElementById('oper').value = 'Del';
  			document.root.submit();
  		}
  	}

  	function DelFolder(file) {
  		var msg = "Favor confirmar que o diretório a seguir será deletado:\n" + document.getElementById('folder').value + "\\" + file.name;
  		var res = confirm(msg);
  		if (res == true) {
  			document.getElementById('file').value = file.name;
  			document.getElementById('oper').value = 'Dlf';
  			document.root.submit();
  		}
  	}

  	function NewFolder(file) {
  		var dir = prompt("Nome do diretório a ser incluído nesta pasta:", "");
  		if (dir != null) {
  			document.getElementById('file').value = dir;
  			document.getElementById('oper').value = 'New';
  			document.root.submit();
  		}
  	}

  	// after the user selects the file they want to upload, submit the form
  	$('#upldFile').on("change", function () {
  		$('#root').submit();
  	});

  </script>

  <style>
	a {
		color: black;
		text-decoration: none;
	}
	a:hover {
		color:black;
		text-decoration: none;
	}

	img {
		border-style: none;
	}
	
	/* hide the file input. important to position it offscreen as opposed display:none. some browsers don't like that */
	#upldFile { position: absolute; left: -9999em; }

	/* an example of styling your label to look like a button */
	#button 
	{
		display: block;
		width: 24px;
		height: 24px;
		text-indent: -9999em;
		background: transparent url(img/icons/glyphicons_181_download_alt.png) 0 0 no-repeat;
	}

	#button:hover {
		cursor: pointer;
	}
	
  </style>

</head>
<body>
<table style="text-align: left;" border="0" cellpadding="2"
 cellspacing="2">
  <tbody>
    <tr>
      <td style="vertical-align: top;"><a href="/"><img
 style="width: 30px; height: 24px;" alt="" noborder
 src="img/icons/glyphicons_118_embed_close.png"></a></td>
      <td style="vertical-align: middle;"><span
 style="font-family: Courier New;">Advanced IntranetSAR Explorer v<%=version %></span></td>
    </tr>
  </tbody>
</table>
<br>
<form action="root.asp" method="post" enctype="multipart/form-data"
 name="root" id="root" onsubmit="return Validate();">
  <input name="oper" id="oper" type="hidden" value="Shw">
  <input name="folder" id="folder" type="hidden" value="<%=folder %>">
  <input name="file" id="file" type="hidden">
  <table border="0" cellpadding="0" cellspacing="0">
    <tbody><%
	Call ShowPath(folder) %>
    </tbody>
  </table>
</form>
</body>
</html>
