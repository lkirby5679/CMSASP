<%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2006 Dogg Software All Rights Reserved
'<>
'<> By using this program, you are agreeing to the terms of the
'<> SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ------------------------------------------------------------------------------
'	Modifird for SkyPortal by Tom Nance (AKA - SkyDogg)
'	URL:		http://www.skyportal.net
'	Date:		2005
' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------
'	Original Author:		Lewis Moten
'	Date:					March 19, 2002
' ------------------------------------------------------------------------------
dim objUpload,remotePathMapped
dim filename, sString, uString, grpsAllowed, bHasGrpAccess, max, upCntr
redim arrUplds(1,1)
  arrUplds(0,0) = false
  arrUplds(0,1) = ""
bHasGrpAccess = false
sString = ""
uString = ""
filename = ""
'####################################
Class clsUpload
Private mbinData
Private mlngChunkIndex
Private mlngBytesReceived
Private mstrDelimiter
Private CR
Private LF
Private CRLF
Private mobjFieldAry()
Private mlngCount

Private Sub RequestData
Dim llngLength
mlngBytesReceived = Request.TotalBytes
mbinData = Request.BinaryRead(mlngBytesReceived)
End Sub

Private Sub ParseDelimiter()
mstrDelimiter = MidB(mbinData, 1, InStrB(1, mbinData, CRLF) - 1)
End Sub

Private Sub ParseData()
Dim llngStart
Dim llngLength
Dim llngEnd
Dim lbinChunk
llngStart = 1
llngStart = InStrB(llngStart, mbinData, mstrDelimiter & CRLF)
While Not llngStart = 0
llngEnd = InStrB(llngStart + 1, mbinData, mstrDelimiter) - 2
llngLength = llngEnd - llngStart
lbinChunk = MidB(mbinData, llngStart, llngLength)
Call ParseChunk(lbinChunk)
llngStart = InStrB(llngStart + 1, mbinData, mstrDelimiter & CRLF)
Wend
End Sub

Private Sub ParseChunk(ByRef pbinChunk)
Dim lstrName
Dim lstrFileName
Dim lstrContentType
Dim lbinData
Dim lstrDisposition
Dim lstrValue
lstrDisposition = ParseDisposition(pbinChunk)
lstrName = ParseName(lstrDisposition)
lstrFileName = ParseFileName(lstrDisposition)
lstrContentType = ParseContentType(pbinChunk)
If lstrContentType = "" Then
lstrValue = CStrU(ParseBinaryData(pbinChunk))
Else
lbinData = ParseBinaryData(pbinChunk)
End If
Call AddField(lstrName, lstrFileName, lstrContentType, lstrValue, lbinData)
End Sub

Private Sub AddField(ByRef pstrName, ByRef pstrFileName, ByRef pstrContentType, ByRef pstrValue, ByRef pbinData)
Dim lobjField
ReDim Preserve mobjFieldAry(mlngCount)
Set lobjField = New clsField
lobjField.Name = pstrName
lobjField.FilePath = pstrFileName				
lobjField.ContentType = pstrContentType
If LenB(pbinData) = 0 Then
lobjField.BinaryData = ChrB(0)
lobjField.Value = pstrValue
lobjField.Length = Len(pstrValue)
Else
lobjField.BinaryData = pbinData
lobjField.Length = LenB(pbinData)
lobjField.Value = ""
End If
Set mobjFieldAry(mlngCount) = lobjField
mlngCount = mlngCount + 1
End Sub

Private Function ParseBinaryData(ByRef pbinChunk)
Dim llngStart
llngStart = InStrB(1, pbinChunk, CRLF & CRLF)
If llngStart = 0 Then Exit Function
llngStart = llngStart + 4
ParseBinaryData = MidB(pbinChunk, llngStart)
End Function

Private Function ParseContentType(ByRef pbinChunk)
Dim llngStart
Dim llngEnd
Dim llngLength
llngStart = InStrB(1, pbinChunk, CRLF & CStrB("Content-Type:"), vbTextCompare)
If llngStart = 0 Then Exit Function
llngEnd = InStrB(llngStart + 15, pbinChunk, CR)
If llngEnd = 0 Then Exit Function
llngStart = llngStart + 15
If llngStart >= llngEnd Then Exit Function
llngLength = llngEnd - llngStart
ParseContentType = Trim(CStrU(MidB(pbinChunk, llngStart, llngLength)))
End Function

Private Function ParseDisposition(ByRef pbinChunk)
Dim llngStart
Dim llngEnd
Dim llngLength
llngStart = InStrB(1, pbinChunk, CRLF & CStrB("Content-Disposition:"), vbTextCompare)
If llngStart = 0 Then Exit Function
llngEnd = InStrB(llngStart + 22, pbinChunk, CRLF)
If llngEnd = 0 Then Exit Function
llngStart = llngStart + 22
If llngStart >= llngEnd Then Exit Function
llngLength = llngEnd - llngStart
ParseDisposition = CStrU(MidB(pbinChunk, llngStart, llngLength))
End Function

Private Function ParseName(ByRef pstrDisposition)
Dim llngStart
Dim llngEnd
Dim llngLength
llngStart = InStr(1, pstrDisposition, "name=""", vbTextCompare)
If llngStart = 0 Then Exit Function
llngEnd = InStr(llngStart + 6, pstrDisposition, """")
If llngEnd = 0 Then Exit Function
llngStart = llngStart + 6
If llngStart >= llngEnd Then Exit Function
llngLength = llngEnd - llngStart
ParseName = Mid(pstrDisposition, llngStart, llngLength)
End Function
' ------------------------------------------------------------------------------
Private Function ParseFileName(ByRef pstrDisposition)
Dim llngStart
Dim llngEnd
Dim llngLength
llngStart = InStr(1, pstrDisposition, "filename=""", vbTextCompare)
If llngStart = 0 Then Exit Function
llngEnd = InStr(llngStart + 10, pstrDisposition, """")
If llngEnd = 0 Then Exit Function
llngStart = llngStart + 10
If llngStart >= llngEnd Then Exit Function
llngLength = llngEnd - llngStart
ParseFileName = Mid(pstrDisposition, llngStart, llngLength)
End Function

Public Property Get Count()
Count = mlngCount
End Property

Public Default Property Get Fields(ByVal pstrName)
Dim llngIndex
If IsNumeric(pstrName) Then
llngIndex = CLng(pstrName)
If llngIndex > mlngCount - 1 Or llngIndex < 0 Then
Call Err.Raise(vbObjectError + 1, "inc_clsUpload.asp", "Object does not exist within the ordinal reference.")
Exit Property
End If
Set Fields = mobjFieldAry(pstrName)
Else
pstrName = LCase(pstrname)
For llngIndex = 0 To mlngCount - 1
If LCase(mobjFieldAry(llngIndex).Name) = pstrName Then
Set Fields = mobjFieldAry(llngIndex)
Exit Property
End If
Next
End If
Set Fields = New clsField
End Property

Private Sub Class_Terminate()
Dim llngIndex
For llngIndex = 0 To mlngCount - 1
Set mobjFieldAry(llngIndex) = Nothing

Next
ReDim mobjFieldAry(-1)
End Sub

Private Sub Class_Initialize()
ReDim mobjFieldAry(-1)
CR = ChrB(Asc(vbCr))
LF = ChrB(Asc(vbLf))
CRLF = CR & LF
mlngCount = 0
Call RequestData
Call ParseDelimiter()
Call ParseData
End Sub

Private Function CStrU(ByRef pstrANSI)
Dim llngLength
Dim llngIndex
llngLength = LenB(pstrANSI)
For llngIndex = 1 To llngLength
CStrU = CStrU & Chr(AscB(MidB(pstrANSI, llngIndex, 1)))
Next
End Function

Private Function CStrB(ByRef pstrUnicode)
Dim llngLength
Dim llngIndex
llngLength = Len(pstrUnicode)
For llngIndex = 1 To llngLength
CStrB = CStrB & ChrB(Asc(Mid(pstrUnicode, llngIndex, 1)))
Next
End Function
End Class
'####################################
Class clsField
Public Name
Private mstrPath
Public FileDir
Public FileExt
Public FileName
Public ContentType
Public Value
Public BinaryData
Public Length
Private mstrText

Public Property Get BLOB()
BLOB = BinaryData
End Property

Public Function BinaryAsText()
Dim lbinBytes
Dim lobjRs
If Length = 0 Then Exit Function
If LenB(BinaryData) = 0 Then Exit Function

If Not Len(mstrText) = 0 Then
BinaryAsText = mstrText
Exit Function
End If
lbinBytes = ASCII2Bytes(BinaryData)
mstrText = Bytes2Unicode(lbinBytes)
BinaryAsText = mstrText
End Function

Public Sub SaveAs(ByRef pstrFileName)
Const adTypeBinary=1
Const adSaveCreateOverWrite=2
Dim lobjStream
Dim lobjRs
Dim lbinBytes
'check length
If Length = 0 Then Exit Sub
'check size
If LenB(BinaryData) = 0 Then Exit Sub

Set lobjStream = Server.CreateObject("ADODB.Stream")
lobjStream.Type = adTypeBinary
Call lobjStream.Open()
lbinBytes = ASCII2Bytes(BinaryData)
Call lobjStream.Write(lbinBytes)

On Error Resume Next

Call lobjStream.SaveToFile(pstrFileName, adSaveCreateOverWrite)

if err<>0 then response.Write "<br>"&err.Description

Call lobjStream.Close()
Set lobjStream = Nothing
End Sub

Public Property Let FilePath(ByRef pstrPath)
mstrPath = pstrPath
If Not InStrRev(pstrPath, ".") = 0 Then
FileExt = Mid(pstrPath, InStrRev(pstrPath, ".") + 1)
FileExt = UCase(FileExt)
End If
If InStrRev(pstrPath, "\") = 0 Then
 FileName=pstrPath
Else
 FileName = Mid(pstrPath, InStrRev(pstrPath, "\") + 1)
End If
If Not InStrRev(pstrPath, "\") = 0 Then
FileDir = Mid(pstrPath, 1, InStrRev(pstrPath, "\") - 1)
End If
End Property

Public Property Get FilePath()
FilePath = mstrPath
End Property

private Function ASCII2Bytes(ByRef pbinBinaryData)
Const adLongVarBinary=205
Dim lobjRs
Dim llngLength
Dim lbinBuffer
llngLength = LenB(pbinBinaryData)
Set lobjRs = Server.CreateObject("ADODB.Recordset")
Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
Call lobjRs.Open()
Call lobjRs.AddNew()
Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0))
Call lobjRs.Update()
lbinBuffer = lobjRs.Fields("BinaryData").GetChunk(llngLength)
Call lobjRs.Close()
Set lobjRs = Nothing
ASCII2Bytes = lbinBuffer
End Function

Private Function Bytes2Unicode(ByRef pbinBytes)
Dim lobjRs
Dim llngLength
Dim lstrBuffer
llngLength = LenB(pbinBytes)
Set lobjRs = Server.CreateObject("ADODB.Recordset")
Call lobjRs.Fields.Append("BinaryData", adLongVarChar, llngLength)
Call lobjRs.Open()
Call lobjRs.AddNew()
Call lobjRs.Fields("BinaryData").AppendChunk(pbinBytes)
Call lobjRs.Update()
lstrBuffer = lobjRs.Fields("BinaryData").Value
Call lobjRs.Close()
Set lobjRs = Nothing
Bytes2Unicode = lstrBuffer
End Function
End Class
'######################################################################################

':: get image size and dimensions
function GetBytes(flnm, offset, bytes)
     Dim obFSO
     Dim obFTemp
     Dim obTextStream
     Dim lngSize
     on error resume next
     Set obFSO = CreateObject("Scripting.FileSystemObject")
     ' First, we get the filesize
     Set obFTemp = obFSO.GetFile(flnm)
     lngSize = obFTemp.Size
     set obFTemp = nothing

     fsoForReading = 1
     Set obTextStream = obFSO.OpenTextFile(flnm, fsoForReading)
     if offset > 0 then
        strBuff = obTextStream.Read(offset - 1)
     end if
     if bytes = -1 then		' Get All!
        GetBytes = obTextStream.Read(lngSize)  'ReadAll
     else
        GetBytes = obTextStream.Read(bytes)
     end if
     obTextStream.Close
     set obTextStream = nothing
     set obFSO = nothing
end function

function lngConvert(strTemp)
     lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
end function

function lngConvert2(strTemp)
     lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
end function

function imgSizeChk(flnm, width, height, depth, strImageType)
     dim strPNG 
     dim strGIF
     dim strBMP
     dim strType
     strType = ""
     strImageType = "(unknown)"
     imgSizeChk = False
     strPNG = chr(137) & chr(80) & chr(78)
     strGIF = "GIF"
     strBMP = chr(66) & chr(77)
     strType = GetBytes(flnm, 0, 3)
     if strType = strGIF then				' is GIF
        strImageType = "GIF"
        Width = lngConvert(GetBytes(flnm, 7, 2))
        Height = lngConvert(GetBytes(flnm, 9, 2))
        Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
        imgSizeChk = True
     elseif left(strType, 2) = strBMP then		' is BMP
        strImageType = "BMP"
        Width = lngConvert(GetBytes(flnm, 19, 2))
        Height = lngConvert(GetBytes(flnm, 23, 2))
        Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
        imgSizeChk = True
     elseif strType = strPNG then			' Is PNG
        strImageType = "PNG"
        Width = lngConvert2(GetBytes(flnm, 19, 2))
        Height = lngConvert2(GetBytes(flnm, 23, 2))
        Depth = getBytes(flnm, 25, 2)
        select case asc(right(Depth,1))
           case 0
              Depth = 2 ^ (asc(left(Depth, 1)))
              imgSizeChk = True
           case 2
              Depth = 2 ^ (asc(left(Depth, 1)) * 3)
              imgSizeChk = True
           case 3
              Depth = 2 ^ (asc(left(Depth, 1)))  '8
              imgSizeChk = True
           case 4
              Depth = 2 ^ (asc(left(Depth, 1)) * 2)
              imgSizeChk = True
           case 6
              Depth = 2 ^ (asc(left(Depth, 1)) * 4)
              imgSizeChk = True
           case else
              Depth = -1
        end select
     else
        strBuff = GetBytes(flnm, 0, -1)		' Get all bytes from file
        lngSize = len(strBuff)
        flgFound = 0
        strTarget = chr(255) & chr(216) & chr(255)
        flgFound = instr(strBuff, strTarget)
        if flgFound = 0 then
           exit function
        end if
        strImageType = "JPG"
        lngPos = flgFound + 2
        ExitLoop = false
		
        do while ExitLoop = False and lngPos < lngSize
           do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
              lngPos = lngPos + 1
           loop
           if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
              lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
              lngPos = lngPos + lngMarkerSize  + 1
           else
              ExitLoop = True
           end if
       loop
       '
       if ExitLoop = False then
          Width = -1
          Height = -1
          Depth = -1
       else
          Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
          Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
          Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
          imgSizeChk = True
       end if
     end if
end function
':: end picture dimension functions :::

Function logActivity(txtToLog)
  if FSOenabled then
'on error resume next
    'log
    if logFlag = "1" then
        if logFile = "" then
           logFile = "upload.txt "
        end if   
	  Set fsoLog = Server.CreateObject("Scripting.FileSystemObject")        
	  Set logFile = fsoLog.OpenTextFile(remotePathMapped & "\" & logFile, 8, True)
	  logFile.WriteLine(txtToLog)
	  logFile.close
	  set logFile = nothing
	  set fsoLog = nothing
    end if
  end if
end function

function checkExt(byRef sName, byRef sExt)
  dim allowed, upl
  allowed = false
  if ar = true then
   for upl = 0 to ubound(extAllowed)
  	if lcase(sExt) = lcase(extAllowed(upl)) then
  	  allowed = true
  	end if	
   next
  else
   if lcase(extAllowed) = lcase(sExt) then
  	allowed = true
   end if  
  end if
  
  if allowed = false then
	sString = sString & "<li>" & txtFileNotAllowed & " - <b>." & sExt & "</b></li>"
	'log
	txt = txtDate & ": " & Date() & "- " & txtAction & ": " & txtBadFileType & "(" & sExt & ") - " & txtUsrName & ": " & session.contents("loggedUser") & " - " & txtFileName & ": " & sName & " - " & txtUploaded & ": " & txtNo
	logActivity(txt)
  end if
  checkExt = allowed
end function

function checkSize(byRef sName, byRef sSize)
  dim allowed
  	allowed = false
  if sSize > sizeLimit then
  	allowed = false
  else
  	allowed = true 
  end if
  
  if allowed = false then
	sString = sString & "<li>" & txtFileTooLg & " '<b>" & (sizeLimit/1000) & " kb</b>'</li>" 
	sString = sString & "<li>" & txtFileSzIs & " '<b>" & FormatNumber(sSize/1000,0) & " kb</b>'</li>"
	'log
	txt = txtDate & ": " & Date & "- " & txtAction & ": " & txtBadFileSize & "(" & FormatNumber(sSize/1000,0) & " kb) - " & txtUsrName & ": " & session.contents("loggedUser") & " - " & txtFileName & ": " & sName & " - " & txtUploaded & ": " & txtNo
	logActivity(txt)
  end if
  checkSize = allowed
end function

function checkThere(byRef sName, byRef sSize)
  dim allowed
  if sSize > 0 then
   if sName <> "" then
  	allowed = true
   else
  	allowed = false 
   end if
  else
  	allowed = false
  end if
  
  if allowed = false then
	sString = sString & "<li>" & txtNoFile & "</li>"
  end if
  checkThere = allowed
end function

function DateToStr3(dtDateTime)
	DateToStr3 = year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime)) & ""
end function

function doublenum(fNum)
	if fNum > 9 then 
		doublenum = fNum 
	else 
		doublenum = "0" & fNum
	end if
end function

function addslash(path)
if right(path,1)="\" then addslash=path else addslash=path & "\"
end function

sub Upload()
  dim f,i,name,path,size,success,memID,ext

  set objUpload=New clsUpload

  success=false
  'targetPath=objUpload.Fields("folder").Value
  max=objUpload.Fields("max").Value
  if max = "" or max < 1 then 
    max = 1
  end if
  upCntr = 0
  memID = objUpload.Fields("memID").Value
  today = datetostr3(now())

'if hasAccess(grpsAllowed) then
 bHasGrpAccess = true
 for i = 1 to max
  name=objUpload.Fields("file" & i).FileName
  filename=objUpload.Fields("file" & i).FileName
  size=objUpload.Fields("file" & i).Length
  ext = objUpload.Fields("file" & i).FileExt

 if checkThere(name,size)=true and checkExt(name,ext)=true and checkSize(name,size)=true then 
  upCntr = upCntr + 1
  uploadPg = true
  redim preserve arrUplds(upCntr,1)
  arrUplds(0,0) = true
  arrUplds(0,1) = ""
  arrUplds(upCntr,0) = filename
  
  'build the full path name
  filename = today & "_" & memID & "_" & i & "." & ext
  path=addslash(remotePathMapped) & filename
  'this line tells it to upload.
  objUpload.Fields("file" & i).SaveAs path
  
  arrUplds(upCntr,1) = filename
  
  'check to validate the upload
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  if objFSO.FileExists(path) then
	'on error resume next
	set f=objFSO.GetFile(path)
	if IsObject(f) then
	  if f.Size=size then success=true else success=false
	end if
    set f=nothing
  end if

  'if upload is an image, attempt to resize if needed and create a thumbnail.
   if (lcase(ext) = "gif" or lcase(ext) = "jpg" or lcase(ext) = "png" or lcase(ext) = "bmp") and success = true and intResize = 1 then '
	if imgSizeChk(path, w, h, c, strType) = true then
      if h > intMaxH or w > intMaxW then
	   ' response.Write("remotePath:" & remotePath & "<br>")
	    'response.Write("remotePath mapped:" & server.MapPath(remotePath) & "<br>")
	    'response.Write("intMaxW:" & intMaxW & "<br>")
	    'response.Write("intMaxH:" & intMaxH & "<br>")
	    'response.Write("filename:" & filename & "<br>")
	     'response.Write("Start resize<br>")
		    ResizeUploadedFiles remotePath,"_rs",intMaxW,intMaxH,rsQuality,false,filename
	     'response.Write("finish resize<br>")
	  else
	    'rename
		Old_ext = lcase("."&ext&"")
		new_ext = lcase("_rs."&ext&"")
		copyTo = replace(lcase(path),Old_ext,new_ext)
		objFSO.CopyFile path,copyTo
      end if
	  if intDoThumb = 1 then
	     'response.Write("<br>Start make thumb<br>")
        if h > intMaxTH or w > intMaxTW then
	     'response.Write("Start thumb resize<br>")
          ResizeUploadedFiles remotePath,"_sm",intMaxTW,intMaxTH,rsQualityThumb,false,filename
	     'response.Write("finish thumb resize<br>")
	    else
	     'response.Write("No resize needed<br>")
	      'rename
		  Old_ext = lcase("."&ext&"")
		  new_ext = lcase("_sm."&ext&"")
		  copyTo = replace(lcase(path),Old_ext,new_ext)
		  objFSO.CopyFile path,copyTo
        end if
	  end if
    end if
   end if
  set objFSO = nothing
  if not success then
    uString = uString & "<li><span class=""fAlert"">failed</span></li>"
  end if
 end if
 next
'end if ':: hasAccess() check
'response.write "<br>" & w & " x " & h & " " & c & " colors"
'response.End()
end sub
'###########################################################################
function moveToLoc(loc)
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(loc) & "\"
		if fso.FolderExists(server.MapPath(loc)) = false then
			fso.CreateFolder(server.MapPath(loc))
		end if
		if fso.FolderExists(server.MapPath(loc)) = false then
			sString = sString & "<li>" & loc & " " & txtNotCreated & "</li>"
		end if
		if fso.FolderExists(dirPath & parent) = false and sString = "" then
			fso.CreateFolder(dirPath & parent)
		end if
		if fso.FolderExists(dirPath & parent) = false and sString = "" then
			sString = sString & "<li>" & loc & parent & "<br>" & txtNotCreated & "</li>"
		end if
		if fso.FolderExists(dirPath & parent & "\" & cat) = false and sString = "" then
			fso.CreateFolder(dirPath & parent & "\" & cat)
		end if
		if fso.FolderExists(dirPath & parent & "\" & cat) = false and sString = "" then
			sString = sString & "<li>" & loc & "\" & parent & "\" & cat & "<br>" & txtNotCreated & "</li>"
		end if
		if fso.FileExists(dirPath & uLoad) = true then
			fso.MoveFile dirPath & uLoad, dirPath & parent & "\" & cat & "\" & uLoad
		else
			'sString = sString & "<li>Failed to Upload file</li>"
		end if
		if not fso.FileExists(dirPath & parent & "\" & cat & "\" & uLoad) = true then
		  if sString = "" then
			sString = sString & "<li>" & txtFileNoMove & "</li>"
		  end if
		end if
	set fso = nothing
end function

function chkIsFileThere(daPath)
  isThere = false
  if FSOenabled then
    set obFSO = Server.CreateObject("Scripting.FileSystemObject")
	  if obFSO.FileExists(daPath) = true then
	    isThere = true
	  end if
	set obFSO = nothing
  end if
  chkIsFileThere = isThere
end function

function deleteFile(daPath)
  if FSOenabled then
    set obFSO = Server.CreateObject("Scripting.FileSystemObject")
	  if obFSO.FileExists(daPath) = true then
	    obFSO.DeleteFile(daPath)
	  end if
	set obFSO = nothing
  end if
end function

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '''''' START UPLOAD CONFIG '''''''''''''''''''''''''''''
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If session.Contents("uploadType") <> "" Then
	'response.Write("hello: <br>")
  'If session.Contents("uploadType") = "xxxxx" Then
  	uploadType = session.Contents("uploadType")
	
	'on error resume next
	set myX_Conn = Server.CreateObject("ADODB.Connection")
	myX_Conn.Errors.Clear
	myX_Conn.Open strConnString
	
	sSQL = "select * from " & strTablePrefix & "UPLOAD_CONFIG where ID = " & uploadType
	set rsU = myX_Conn.execute(sSQL)
	remotePath = rsU("UP_FOLDER")
	'response.Write("remotePath: " & remotePath)
  	remotePathMapped = Server.MapPath(remotePath)
  	'sizeLimit = 10000000			         
  	sizeLimit = (rsU("UP_SIZELIMIT")*1000) 'maximum size (bytes) of the file to be uploaded
	tmpAllow = rsU("UP_ALLOWEDEXT")
  	if instr(tmpAllow,",") then
  		extAllowed = split(tmpAllow,",")
		ar = true
  	else
  		extAllowed = chkString(tmpAllow,"")
		ar = false
  	end if
 	' extAllowed = Array("gif", "jpg")
  	logFlag = rsU("UP_LOGUSERS")     ' 1 = logs the upload activity, 0 = doesn't
  	logFile = rsU("UP_LOGFILE")      'this is the file that logs all the upload activity.
  	grpsAllowed = rsU("UP_ALLOWEDGROUPS")
	active = rsU("UP_ACTIVE")
	intMaxTW = cint(rsU("UP_THUMB_MAX_W"))
	intMaxTH = cint(rsU("UP_THUMB_MAX_H"))
	intMaxW = cint(rsU("UP_NORM_MAX_W"))
	intMaxH = cint(rsU("UP_NORM_MAX_H"))
	intResize = cint(rsU("UP_RESIZE"))
	intDoThumb = cint(rsU("UP_CREATE_THUMB")) 
	
	if intResize = 1 then
	Dim max_Quality, max_QualityThumb
    rsQuality = "90"
    rsQualityThumb = "70"
    end if
    
	set rsU = nothing
  	myX_Conn.close
	set myX_Conn = nothing
	
	  upload()
	'response.End()
  end if
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '''''' END UPLOAD CONFIG '''''''''''''''''''''''''''''
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">

Sub ResizeUploadedFiles(max_path, max_Suffix, max_maxWidth, max_maxHeight, max_Quality, max_RemoveOrig, up_filename)
  Dim max_keys, max_i, max_curKey, max_fileName, max_fso, max_newFileName, max_curPath, max_curName, max_curExt, max_lastPos, max_orgCurPath
  if max_path <> "" and right(max_path,1) <> "/" then max_path = max_path & "/"
  Set max_fso = CreateObject("Scripting.FileSystemObject")  
  max_maxWidth = Cint(max_maxWidth)
  max_maxHeight  = Cint(max_maxHeight)  
          max_fileName = up_filename
          if max_fileName <> "" then
            max_curPath = "" : max_curName = "" : max_curExt = ""
            max_lastPos = InStrRev(max_fileName,"/")
            if max_lastPos > 0 then
              max_curPath = mid(max_fileName,1,max_lastPos)	
              max_curName = mid(max_fileName,max_lastPos+1,Len(max_fileName)-max_lastPos)	
              max_fileName = up_filename            
            else
              max_curName = up_filename	
            end if
            max_lastPos = InStrRev(max_curName,".")
            if max_lastPos > 0 then
              max_curExt = mid(max_curName,max_lastPos+1,Len(max_curName)-max_lastPos)	
              max_curName = mid(max_curName,1,max_lastPos-1)
            end if
            max_curExt = LCase(max_curExt)
     		max_orgCurPath = max_curPath
            if max_curPath = "" then max_curPath = max_path
	    'response.Write("file exist:" & server.MapPath(max_curPath & up_filename) & "<br>")
            if max_fso.FileExists(Server.MapPath(max_curPath & up_filename)) then
                max_newFileName = max_curName & max_Suffix & "." & max_curExt
	    'response.Write("max_newFileName:" & max_newFileName & "<br>")
                FitImage_Comp "includes/image_resizer.aspx", Server.MapPath(max_CurPath & max_fileName), Server.MapPath(max_curPath & max_newFileName), max_maxWidth, max_maxHeight, max_Quality
                if max_RemoveOrig then
                  if LCase(max_fileName) <> LCase(max_newFileName) then
                    'max_fso.DeleteFile Server.MapPath(max_curPath & max_fileName)
                  end if  
                end if
            end if
          end if	
End Sub

sub FitImage_Comp(DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality)
    select case DetectDotNetComponent(DotNetResize)
    case "DOTNET1"
      Image_Size_DotNet "Msxml2.ServerXMLHTTP.4.0",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality
    case "DOTNET2"
      Image_Size_DotNet "Msxml2.ServerXMLHTTP",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality
    case "DOTNET3"
      Image_Size_DotNet "Microsoft.XMLHTTP",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality
    end select
end sub

function DetectDotNetComponent(DotNetResize)
  Dim DotNetImageComponent, ResizeComUrl, LastPath
	
	DotNetImageComponent = ""
	ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
	LastPath = InStrRev(ResizeComUrl,"/")
	if LastPath > 0 then
		ResizeComUrl = left(ResizeComUrl,Lastpath)
	end if
	ResizeComUrl = ResizeComUrl & DotNetResize
	'Response.Write ResizeComUrl & "<br>"
	
	'Check for ASP.NET 1
	if DotNetCheckComponent("Msxml2.ServerXMLHTTP.4.0", ResizeComUrl) = true then
		'Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP.4.0<br>"
		DotNetImageComponent = "DOTNET1"
	else
		if DotNetCheckComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
			'Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP<br>"
			DotNetImageComponent = "DOTNET2"
		else
			if DotNetCheckComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
				'Response.Write "FOUND: ASP.NET Microsoft.XMLHTTP<br>"
				DotNetImageComponent = "DOTNET3"
			else
				Response.Write "NOT FOUND: ASP.NET Server Component<br>"
			end if
		end if
	end if
	'on error goto 0
  
	DetectDotNetComponent = DotNetImageComponent
end function

function DotNetCheckComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
    objHttp.open "GET", ResizeComUrl, false
	if err.number = 0 then
      objHttp.Send ""
			if (objHttp.status <> 200 ) then
				Response.Write "An error has accured with ASP.NET component " & DotNetObj & "<br>"
				Response.Write "Error:<br>" & objHttp.responseText & "<br>"
				Response.End
			end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detection = true
      end if
	end if
  End if
  Set objHttp = nothing
  'on error goto 0
  DotNetCheckComponent = Detection
end function

sub Image_Size_DotNet(DotNetComp, DotNetResize, imgFile,newImgFile,maxWidth,maxHeight,Quality)
  Dim objHttp, ResizeComUrl, ResizeParams, LastPath
  'Response.Write "Image_Size_DotNet<br>"
  ResizeParams = "?f=" & Server.UrlEncode(imgFile) & "&nf=" & Server.UrlEncode(newImgFile) & "&w=" & maxWidth & "&h=" & maxHeight & "&q=" & Quality
  ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
  LastPath = InStrRev(ResizeComUrl,"/")
  if LastPath > 0 then
    ResizeComUrl = left(ResizeComUrl,Lastpath)
  end if
  ResizeComUrl = ResizeComUrl & DotNetResize & ResizeParams
  'Response.Write ResizeComUrl & "<br>"

  on error resume next
  set objHttp = Server.CreateObject(DotNetComp)
  if err.number <> 0 then
    Response.Write "ERROR: ASP.NET (" & DotNetComp & ") is not installed!<br>Image resize is not available"
    Response.End
  end if
  
  objHttp.open "GET", ResizeComUrl, false
  objHttp.Send ""
  
  ' Check notification validation
  if (objHttp.status <> 200 ) then
    ' HTTP error handling
    Response.Write "HTTP ERROR: " & objHttp.status & "<br>"
    Response.Write "Returned:<br>" & objHttp.responseText 
    
  elseif (objHttp.responseText = "Done") then
  'Response.Write "it says DONE<br>"
  else
    if trim(objHttp.responseText)="" or instr(objHttp.responseText,"@ Page Language=""C#""")>0 then
      Response.Write "DOT NET Unsupported"
	else
  	  'Response.Write "unspecified error: " & objHttp.responseText & "<br>"
    end if
  end if
  Set objHttp = Nothing
  'on error goto 0
end sub

</SCRIPT>