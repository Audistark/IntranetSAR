<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
'Limit of upload size
Dim UploadSizeLimit

'********************************** GetUpload **********************************
'This function reads all form fields from binary input and returns it as a dictionary object.
'The dictionary object containing form fields. Each form field is represented by six values :
'.Name name of the form field (<Input Name="..." Type="File,...">)
'.ContentDisposition = Content-Disposition of the form field
'.FileName = Source file name for <input type=file>
'.ContentType = Content-Type for <input type=file>
'.Value = Binary value of the source field.
'.Length = Len of the binary data field
Function GetUpload()

  Dim Result
  Set Result = Nothing
  '-- Se o form foi submetido por POST
  If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  
    Dim CT, PosB, Boundary, Length, PosE
    CT = Request.ServerVariables("HTTP_Content_Type") 'Tipo do cabeçalho
    If LCase(Left(CT, 19)) = "multipart/form-data" Then 'Content-Type do cabeçalho deve ser "multipart/form-data"

      'This is upload request.
      'Get the boundary and length from Content-Type header
      PosB = InStr(LCase(CT), "boundary=") 'Finds boundary
      If PosB > 0 Then Boundary = Mid(CT, PosB + 9) 'Separetes boundary

      '****** Error of IE5.01 - doubbles http header
      PosB = InStr(LCase(CT), "boundary=") 
      If PosB > 0 then 'Patch for the IE error
        PosB = InStr(Boundary, ",")
        If PosB > 0 Then Boundary = Left(Boundary, PosB - 1)
      end if
      '****** Error of IE5.01 - doubbles http header

      Length = CLng(Request.ServerVariables("HTTP_Content_Length")) 'Get Content-Length header
      If "" & UploadSizeLimit <> "" Then
        UploadSizeLimit = CLng(UploadSizeLimit)
        If Length > UploadSizeLimit Then
          Request.BinaryRead (Length)
          Err.Raise 2, "GetUpload", "Upload size " & FormatNumber(Length, 0) & "B exceeds limit of " & FormatNumber(UploadSizeLimit, 0) & "B"
          Exit Function
        End If
      End If
      
      If Length > 0 And Boundary <> "" Then 'Are there required informations about upload ?
        Boundary = "--" & Boundary
        Dim Head, Binary
        Binary = Request.BinaryRead(Length) 'Reads binary data from client
        
        'Retrieves the upload fields from binary data
        Set Result = SeparateFields(Binary, Boundary)
        Binary = Empty 'Clear variables
      Else
        Err.Raise 10, "GetUpload", "Zero length request ."
      End If
    Else
      Err.Raise 11, "GetUpload", "No file sent."
    End If
  Else
    Err.Raise 1, "GetUpload", "Bad request method."
  End If
  Set GetUpload = Result
End Function

'********************************** SeparateFields **********************************
'This function retrieves the upload fields from binary data and retuns the fields as array
'Binary is safearray ( VT_UI1 | VT_ARRAY ) of all document raw binary data from input.
Function SeparateFields(Binary, Boundary)
  Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary
  Dim Fields
  Boundary = StringToBinary(Boundary)

  PosOpenBoundary = InStrB(Binary, Boundary)
  PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)

  Set Fields = CreateObject("Scripting.Dictionary")
  Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)
    'Header and file/source field data
    Dim HeaderContent, bFieldContent
    'Header fields
    Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type
    'Helping variables
    Dim Field, TwoCharsAfterEndBoundary
    'Get end of header
    PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))

    'Separates field header
    HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
    
    'Separates field content
    bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)

    'Separates header fields from header
    GetHeadFields BinaryToString(HeaderContent), Content_Disposition, FormFieldName, SourceFileName, Content_Type

    'Create one field and assign parameters
    If Len(SourceFileName) > 0 Then
	    Set Field = CreateUploadField()'See the JS function bellow

		Field.ContentDisposition = Content_Disposition
		Field.FilePath = SourceFileName
		Field.FileName = GetFileName(SourceFileName)
		Field.ContentType = Content_Type
		Field.Value = bFieldContent
		Field.Length = LenB(bFieldContent)
	Else
	    Set Field = CreateField()'See the JS function bellow

		Field.Value = BinaryToString(bFieldContent)
		Field.Length = Len(Value)
	End If
    Field.Name = FormFieldName

'	response.write "<br>:" & FormFieldName
    Fields.Add FormFieldName, Field

    'Is this last boundary ?
    TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
    isLastBoundary = TwoCharsAfterEndBoundary = "--"

    If Not isLastBoundary Then 'This is not last boundary - go to next form field.
      PosOpenBoundary = PosCloseBoundary
      PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
    End If
  Loop
  Set SeparateFields = Fields
End Function

'********************************** Utilities **********************************

'Separates header fields from upload header
Function GetHeadFields(ByVal Head, Content_Disposition, Name, FileName, Content_Type)
  Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))

  Name = (SeparateField(Head, "name=", ";")) 'ltrim
  If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)

  FileName = (SeparateField(Head, "filename=", ";")) 'ltrim
  If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)

  Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function

'Separates one field between sStart and sEnd
Function SeparateField(From, ByVal sStart, ByVal sEnd)
  Dim PosB, PosE, sFrom
  sFrom = LCase(From)
  PosB = InStr(sFrom, sStart)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    PosE = InStr(PosB, sFrom, sEnd)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(From, PosB, PosE - PosB)
  Else
    SeparateField = Empty
  End If
End Function

'Separetes file name from the full path of file
Function GetFileName(FullPath)
  Dim Pos, PosF
  PosF = 0
  For Pos = Len(FullPath) To 1 Step -1
    Select Case Mid(FullPath, Pos, 1)
      Case "/", "\": PosF = Pos + 1: Pos = 0
    End Select
  Next
  If PosF = 0 Then PosF = 1
  GetFileName = Mid(FullPath, PosF)
End Function

Function BinaryToString(Binary)
	BinaryToString = RSBinaryToString(Binary)
End Function

Function RSBinaryToString(xBinary)
  'This function converts binary data (VT_UI1 | VT_ARRAY or MultiByte string)
	'to string (BSTR) using ADO recordset
	'The fastest way - requires ADODB.Recordset
	'Use this function instead of BinaryToString if you have ADODB.Recordset installed
	'to eliminate problem with PureASP performance

	Dim Binary
	'MultiByte data must be converted to VT_UI1 | VT_ARRAY first.
	if vartype(xBinary)=8 then Binary = MultiByteToBinary(xBinary) else Binary = xBinary

  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)

	if LBinary>0 then
		RS.Fields.Append "mBinary", adLongVarChar, LBinary
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk Binary
		RS.Update
		RSBinaryToString = RS("mBinary").Value
	Else
		RSBinaryToString = ""
	End If
	Set RS = Nothing
End Function

Function MultiByteToBinary(MultiByte)
  ' This function converts multibyte string to real binary data (VT_UI1 | VT_ARRAY)
  ' Using recordset
  Dim RS, LMultiByte, Binary
  Const adLongVarBinary = 205
  Set RS = CreateObject("ADODB.Recordset")
  LMultiByte = LenB(MultiByte)
	if LMultiByte>0 then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	End If
  MultiByteToBinary = Binary
End Function

Function StringToBinary(String)
  Dim I, B
  For I=1 to len(String)
    B = B & ChrB(Asc(Mid(String,I,1)))
  Next
  StringToBinary = B
End Function

'The function simulates save binary data using conversion to string and filesystemobject
Function vbsSaveAs(FileName, ByteArray)
	Dim FS, TextStream
	Set FS = CreateObject("Scripting.FileSystemObject")

	Set TextStream = FS.CreateTextFile(FileName)
	
	'And this is the problem why only short files - BinaryToString uses byte-to-char VBS conversion. It takes a lot of computer time.
	TextStream.Write BinaryToString(ByteArray) ' BinaryToString is in upload.inc.
	TextStream.Close
End Function

Function RetornaMime(Extensao)
	Dim d                   
	Set d = CreateObject("Scripting.Dictionary")
	d.Add ".doc", "application/msword"
	d.Add ".dot", "application/msword"          
	d.Add ".htm", "text/html"
	d.Add "html", "text/html"
	d.Add ".bmp", "image/bmp"
	d.Add ".gif", "image/gif "
	d.Add ".jpe", "image/jpeg"
	d.Add "jpeg", "image/jpeg"
	d.Add ".jpg", "image/jpeg"
	d.Add ".tif", "image/tiff"
	d.Add "tiff", "image/tiff"
	d.Add ".xla", "application/vnd.ms-excel"
	d.Add ".xlc", "application/vnd.ms-excel"
	d.Add ".xlm", "application/vnd.ms-excel"
	d.Add ".xls", "application/vnd.ms-excel"
	d.Add ".xlt", "application/vnd.ms-excel"
	d.Add ".xlw", "application/vnd.ms-excel"
	d.Add ".pot", "application/vnd.ms-powerpoint"
	d.Add ".pps", "application/vnd.ms-powerpoint"
	d.Add ".ppt", "application/vnd.ms-powerpoint"
	d.Add ".mdb", "application/msaccess"	
	d.Add ".txt", "text/plain"
	d.Add ".pdf", "application/pdf"	
	d.Add ".zip", "application/zip"	
	d.Add ".mpp", "application/vnd.ms-project"
	d.Add ".hlp", "application/winhlp"	
	
	Mime = trim(d.Item(Extensao))	
	IF (Mime =  "") then
		RetornaMime = "application/octet-stream"
	Else
		RetornaMime = Mime
	End if
End Function
</SCRIPT>
<SCRIPT RUNAT=SERVER LANGUAGE=JSCRIPT>
function CreateField(){ return new f_Init() }
function f_Init(){
	this.Name = null;
	this.Value = null;
	this.Length = null;
}

function CreateUploadField(){ return new uf_Init() }
function uf_Init(){
	this.Name = null;
	this.ContentDisposition = null;
	this.FileName = null;
	this.FilePath = null;
	this.ContentType = null;
	this.Value = null;
	this.String = null;
	this.Length = null;
	this.SaveAs = jsSaveAs;
}
function jsSaveAs(FileName){
  return vbsSaveAs(FileName, this.ByteArray)
  
}
</SCRIPT>