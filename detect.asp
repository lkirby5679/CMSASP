<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>SkyPortal component check</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
dim bFSOenabled

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
		Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP.4.0<br>"
		DotNetImageComponent = "DOTNET1"
	else
		if DotNetCheckComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
			Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP<br>"
			DotNetImageComponent = "DOTNET2"
		else
			if DotNetCheckComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
				Response.Write "FOUND: ASP.NET Microsoft.XMLHTTP<br>"
				DotNetImageComponent = "DOTNET3"
			else
				Response.Write "NOT FOUND: ASP.NET Server Component<br>"
			end if
		end if
	end if
	on error goto 0
	
    FSOcomponent = fsoCheck()
    'if FSOcomponent = true then ImageComponent = NetImageComponent  
  
	DetectDotNetComponent = DotNetImageComponent
end function

function DotNetCheckComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
	'response.write("Checking "&DotNetObj&"<br>")
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
  	'response.write("Object "&DotNetObj&" created<br>")
    objHttp.open "GET", ResizeComUrl, false
		if err.number = 0 then
      objHttp.Send ""
			if (objHttp.status <> 200 ) then
				Response.Write "An error has accured with ASP.NET component " & DotNetObj & "<br>"
				Response.Write "Returned:<br>" & objHttp.responseText & "<br>"
				Response.End
			end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detection = true
      end if
		end if
    Set objHttp = nothing
  End if
  on error goto 0
 	'response.write("Detection is "&Detection&"<br>")
  DotNetCheckComponent = Detection
end function

function fsoCheck()
     on error resume next
     err.clear
	 set fso = Server.CreateObject("Scripting.FileSystemObject")
	 if err.number = 0 then
	   Response.Write "FOUND: FileSystemObject scripting component<br>"
	   bFSOenabled = true
	   set fso = nothing
	 else 
	   Response.Write "NOT FOUND: FileSystemObject scripting component<br>"
	   bFSOenabled = false
	 end if
     on error goto 0
	 fsoCheck = bFSOenabled
end function

function getEmailComponents()
Dim arrComponent(10)
Dim arrValue(10)
Dim arrName(10)

' components
arrComponent(0) = "CDO.Message"
arrComponent(1) = "CDONTS.NewMail"
arrComponent(2) = "SMTPsvg.Mailer"
arrComponent(3) = "Persits.MailSender"
arrComponent(4) = "SMTPsvg.Mailer"
arrComponent(5) = "CDONTS.NewMail"
arrComponent(6) = "dkQmail.Qmail"
arrComponent(7) = "Geocel.Mailer"
arrComponent(8) = "iismail.iismail.1"
arrComponent(9) = "Jmail.smtpmail"
arrComponent(10) = "SoftArtisans.SMTPMail"

' component values
arrValue(0) = "cdosys"
arrValue(1) = "cdonts"
arrValue(2) = "aspmail"
arrValue(3) = "aspemail"
arrValue(4) = "aspqmail"
arrValue(5) = "chilicdonts"
arrValue(6) = "dkqmail"
arrValue(7) = "geocel"
arrValue(8) = "iismail"
arrValue(9) = "jmail"
arrValue(10) = "smtp"

' component names
arrName(0) = "CDOSYS (IIS 5/5.1/6)"
arrName(1) = "CDONTS (IIS 3/4/5)"
arrName(2) = "ASPMail"		'yes
arrName(3) = "ASPEMail"	'yes
arrName(4) = "ASPQMail"	'yes			'
arrName(5) = "Chili!Mail (Chili!Soft ASP)"	'
arrName(6) = "dkQMail"						'
arrName(7) = "GeoCel"						'
arrName(8) = "IISMail"					'
arrName(9) = "JMail"						'
arrName(10) = "SA-Smtp Mail"

Response.Write("<ul>") & vbcrlf
'Response.Write("<option value=""none"" selected></option>") & vbcrlf
Dim i
for i=0 to UBound(arrComponent)
	if isInstalled(arrComponent(i)) then
	  Response.Write("<li>"  & arrName(i) &"</li>") & vbcrlf
	end if
next
Response.Write("</ul>") & vbcrlf
end function				'

Function isInstalled(obj)
	on error resume next
	installed = False
	Err = 0
	Dim chkObj
	Set chkObj = Server.CreateObject(obj)
	If 0 = Err Then installed = True
	Set chkObj = Nothing
	isInstalled = installed
	Err = 0
	on error goto 0
End Function

%>
<body>
<h1>SkyPortal component check</h1>
Detecting Components:<br><br>
<% 
det = DetectDotNetComponent("includes/scripts/checkfordotnet.aspx") %>
<h2>
<% if det <> "" and bFSOenabled then %>
  SkyPortal can be fully 
  used on this server!<br><br>
<% else %>
	  SkyPortal's full features are not available.<br><br>
<%
   end if %>
</h2>	  
<%    if det = "" then 
  		response.Write("ASP.NET is NOT installed on this server.<br>")
		response.Write("The image thumbnails will NOT be available.<br><br>")
	  else
  		response.Write("ASP.NET is installed on this server.<br>")
		response.Write("The image resizing will be available.<br><br>")
 	  end if
	  if bFSOenabled = false then
	    response.Write("FileSystemObject is not available on this server<br>")
	    response.Write("Uploads will NOT be available in the installation")
	  else
	    response.Write("FileSystemObject is available on this server<br>")
	    response.Write("Uploads will be available in the installation<br><br>")
	  end if
	    response.Write("<h3>The following compatable email components are available on this server</h3>")
	  getEmailComponents() %>
</body>
</html>
