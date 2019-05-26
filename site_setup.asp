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
itName = "Frosty_Sky"
itFolder = "Frosty_Sky"
itLogo = "site_logo.jpg"
itAuthor = "Frost"
itDesc = "x"
itSubSkin = 1

installTheme = itFolder

'dim arrCntryData(260)
dim inFSOenabled, bErr
%>
<%
Dim strVer
strDebug = false
strVer = "RC3"
longVer = "RC3"
dim fsoMSG
dim fsoObj
dim fso, fo
dim boolPerm
'Dim dbHits
dim arrData()
dim indexes()
Dim portalUrl
Dim newTbl
Dim oldTbl
Dim betaTbl
Dim v20Tbl
Dim v21Tbl
Dim vRC1Tbl
Dim vRC2Tbl
Dim erMsg
dim tmpMSG
'sqlVer = 7
bHasTable = false
sInstallType = ""
bIsUpgrade = false
boolPerm = false
fsoObj = false
fsoMSG = ""
erMsg = ""
dbHits = 0
newTbl = 0
oldTbl = 0
betaTbl = 0
v20Tbl = 0
v21Tbl = 0
vRC1Tbl = 0
vRC2Tbl = 0
comCode = ""
CustomCode = 0

ErrorCount = Request.QueryString("RC")
comCode = cLng(Request.QueryString("cmd"))
sessCode = session.Contents("setup")

if ErrorCount = 1 then
	CustomCode = 2
end if

if comCode <> "3" then 'setup
	blnSetup = "Y"
else ' db is created/updated. 
	blnSetup = ""
end if

if comCode = "3" then ' db is created/updated. 
	blnSetup = ""
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
end if

%>
<!--#INCLUDE FILE="config.asp" -->
<!--#include file="includes/inc_ADOVBS.asp" -->
<!-- #include file="lang/en/core.asp" -->
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="lang/en/core_install_data.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="includes/inc_encryption.asp" -->
<!--#INCLUDE FILE="includes/inc_DBFunctions.asp" -->
<%
installTheme = itFolder
if ErrorCount = 0 then

  portalUrl = "http://" & request.ServerVariables("SERVER_NAME") & left(Request.ServerVariables("URL"),instrrev(Request.ServerVariables("URL"),"/"))
  
if request("debug") = 1 then
  strDebug = true
end if

if comCode = "2" then 
	if strDebug then
	response.Write("<b>comCode = 2, lets get variable info!</b><br><br>")
	end if
			adminName = request.Form("adminName")
			adminPass = pEncrypt(pEnPrefix & Request.Form("adminPass"))
			siteName = replace(ChkString(request.Form("siteName"),"sqlstring"),"'","''")
			mailServer = request.Form("mailServer")
			emailAddy = request.Form("emailAddy")
			emailComponent = request.Form("emailComponent")
			instType = ChkString(request.Form("installType"),"sqlstring")
			localhost = ChkString(request.Form("localhost"),"sqlstring")
			
	'response.Write("<b>instType: </b>" & instType & "<br>")
		if instType <> "new" then
			dim	buArticle, buDL, buClassified, buForums, buPics, buLinks
			buArticle = cint(request.Form("Articles"))
			buDL = cint(request.Form("Downloads"))
			buClassified = cint(request.Form("Classifieds"))
			buForums = cint(request.Form("Forums"))
			buPics = cint(request.Form("Pictures"))
			buLinks = cint(request.Form("Links"))
		end if
	'response.Write("<b>buArticle: </b>" & buArticle & "<br>")
	'response.Write("<b>buDL: </b>" & buDL & "<br>")
	'response.Write("<b>buClassified: </b>" & buClassified & "<br>")
	'response.Write("<b>buForums: </b>" & buForums & "<br>")
	'response.Write("<b>buPics: </b>" & buPics & "<br>")
	'response.Write("<b>buLinks: </b>" & buLinks & "<br>")
	'response.End()
	if strDebug then
	response.Write("Variable listing:<br>")
	response.Write("<b>adminName: </b>" & adminName & "<br>")
	response.Write("<b>adminPass: </b>" & adminPass & "<br>")
	response.Write("<b>siteName: </b>" & siteName & "<br>")
	response.Write("<b>mailServer: </b>" & mailServer & "<br>")
	response.Write("<b>emailAddy: </b>" & emailAddy & "<br>")
	response.Write("<b>emailComponent: </b>" & emailComponent & "<br>")
	response.Write("<b>localhost: </b>" & localhost & "<br>")
	response.Write("<b>installTheme: </b>" & installTheme & "<br><br>")
	end if
end if

	on error resume next
	my_Conn.Errors.Clear
	Err.Clear
	
	'check database type
	Select case lcase(strDBType)
		case "access"
			ErrorCount = 0
		case "sqlserver"
			ErrorCount = 0
	'		CustomCode = 1
		case "mysql"
			ErrorCount = 0
	'		CustomCode = 1
		case else
			ErrorCount = 1
			CustomCode = 1
	end Select
	
	'check if the connection string will open the database
	if ErrorCount = 0 then   'try to open the connection
		set my_Conn = Server.CreateObject("ADODB.Connection")
		my_Conn.Open strConnString	
	
		'if there is an error,  show error box
		for counter = 0 to my_conn.Errors.Count -1
			ConnErrorNumber = my_conn.Errors(counter).Number
			ConnErrorDesc = my_conn.Errors(counter).Description
			if ConnErrorNumber <> 0 then 
				ErrorCount = 1
				CustomCode = 2
				ErrorCode = ConnErrorNumber & "<br>" & ConnErrorDesc
				my_conn.Errors.Clear 
			end if
		next
		my_Conn.Errors.Clear
		Err.Clear
	end if
	' debugging
	if strDebug then
	response.Write("Start FSO check<br>")
	end if
	'check for FileSystemObject
	Err.Clear
	Err = 0
	on error resume next
    set fs = CreateObject("Scripting.FileSystemObject")
	if 0 = Err then
	   fsoObj = true
	end if
	set fs = nothing
	Err.Clear
	' debugging
	if strDebug then
	response.Write("End FSO check: " & fsoObj & "<br><br>")
	end if
	
	'response.Write("<br>bHasTable0:" & bHasTable)
	'test for v1.3x member table
	' if this table is missing, it is a new install
	strSql = "SELECT MEMBER_ID FROM PORTAL_MEMBERS"
	my_Conn.Execute strSql
	Call CheckSqlError("new")
	my_Conn.Errors.Clear
	Err.Clear
	
	'response.Write("<br>bHasTable1:" & bHasTable)
  if bHasTable then
	'Lets test for v1.5 new table field
	'this field is added in v1.5 - not in v1.3x
	strSql = "SELECT B_NAME FROM PORTAL_BANNERS"
	my_Conn.Execute strSql
	Call CheckSqlError("v13x")
	my_Conn.Errors.Clear
	Err.Clear
	
	'response.Write("<br>bHasTable2:" & bHasTable)
	if bHasTable and sInstallType = "v13x" then
	'Lets test for v1.5 default URL in case the folder was renamed
	strSql = "SELECT C_STRHOMEURL FROM PORTAL_CONFIG"
	my_Conn.Execute strSql
	Call CheckSqlError("new")
	'Insert new values
	strSql = "UPDATE PORTAL_CONFIG SET C_STRHOMEURL='" & portalUrl & "' WHERE CONFIG_ID = 1"
	my_Conn.Execute strSql
	Call CheckSqlError("new")
	my_Conn.Errors.Clear
	Err.Clear
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
	end if
  end if
	
	'response.Write("<br>bHasTable3:" & bHasTable)
  if bHasTable then
	'Lets test for v1.5b3 new table field
	'this field is added in v1.5b3 - not in v1.5
	' upgrades to v2.0 either started with v1.3x OR v1.5b3
	strSql = "SELECT THEME_ID FROM PORTAL_MEMBERS"
	my_Conn.Execute strSql
	Call CheckSqlError("v15b3")
	my_Conn.Errors.Clear
	Err.Clear
  end if
	'response.Write("<br>bHasTable4:" & bHasTable)
  if bHasTable then
	'Lets test for v2.1 new table field
	'this field is added in v2.1 - is not in v2.0
	strSql = "SELECT C_SECIMAGE FROM PORTAL_CONFIG"
	my_Conn.Execute strSql
	Call CheckSqlError("v20")
	my_Conn.Errors.Clear
	Err.Clear
  end if
	
	'response.Write("<br>bHasTable5:" & bHasTable)
  if bHasTable then
	'Lets test for SP RC1 new table field
	'this field is added in SP RC1 - is not in v2.1x
	strSql = "SELECT C_INTSUBSKIN FROM PORTAL_CONFIG"
	my_Conn.Execute strSql
	Call CheckSqlError("v21")
	my_Conn.Errors.Clear
	Err.Clear
  end if
  
  if bHasTable then
	'Lets test for SP RC2 new table field
	'this field is added in SP RC2 - is not in RC1
	strSql = "SELECT APP_GROUPS_FULL FROM PORTAL_APPS"
	my_Conn.Execute strSql
	Call CheckSqlError("vRC1")
	my_Conn.Errors.Clear
	Err.Clear
  end if
  
  if bHasTable then
	'Lets test for SP RC3 new table field
	'this field is added in SP RC3 - is not in RC2
	strSql = "SELECT id FROM Menu"
	my_Conn.Execute strSql
	Call CheckSqlError("vRC2")
	my_Conn.Errors.Clear
	Err.Clear
  end if		
	'response.Write("<br>bHasTable6:" & bHasTable)			
	
	on error goto 0
	
end if 'responsecode = 0
Response.Buffer = True
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html> 
<head> 
<!-- This page is generated by SkyPortal / SkyPortal.net <%= date() %> -->
<title>SkyPortal v<%= strVer %> | Site Setup</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<META HTTP-EQUIV="imagetoolbar" CONTENT="no">
<META NAME="AUTHOR" CONTENT="SkyPortal www.SkyPortal.net">
<META NAME="GENERATOR" CONTENT="SkyPortal - http://www.SkyPortal.net">
<meta name="COPYRIGHT" content="Portal code is Copyright (C)2005 - 2006 Tom Nance All Rights Reserved">
<meta http-equiv="Content-Style-Type" content="text/css">
<link rel="stylesheet" href="Themes/<%= installTheme %>/style_core.css" type="text/css">
</head>
<body bgcolor="#C6C9D1" background="Themes/<%= installTheme %>/background.gif" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<script type="text/JavaScript">
function showMe(id) {
	var o1=document.getElementById(id);
		o1.style.display="block";
}
function hideMe(id) {
	var o1=document.getElementById(id);
		o1.style.display="none";
}
</script>

<form action="site_setup.asp?cmd=2" method="post" id="form2" name="form2">
<table class="spThemePage" width="100%" align="center" border="0" cellpadding="0" cellspacing="0"><tr><td>

<a name="top"></a>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
          <td width="300" align="left" valign="middle" class="sp_Header_Left"><a href="default.asp"><img alt="SkyPortal v<%= strVer %>" border="0" src="Themes/<%= installTheme %>/site_Logo.jpg"></a></td>
          <td align="center" class="sp_Header_Tile"><img src="files/banners/SkyPortal.gif" width="389" height="56"> 
          </td>
          <td align="center" class="sp_Header_Rite">
          </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td class="sp_NavLeft">&nbsp;</td>
<td class="sp_NavTile" align="center" width="100%">&nbsp;</td>
<td class="sp_NavRite">&nbsp;</td>
</tr>
</table>

<table class="spThemePage" border="0" width="100%" align="center">
<tr>
<td class="leftPgCol"></td>
<td valign="top" align="center" class="mainPgCol">
<div class="spThemeBlock1">
<div class="spThemeBlock1_tr">
<div class="spThemeBlock1_tl">
<div class="spThemeBlock1_tc">
<h4>SkyPortal v<%= strVer %><font color="#CC0000">&nbsp;<%= txtSUSiteSetup %></font></h4>
</div>
</div>
</div>
<div class="spThemeBlock1_ml">
<div class="spThemeBlock1_mr">
<div class="spThemeBlock1_content">
<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0">

<tr><td bgcolor="#E7E7EA" align="center" valign="top">
<% 
	
	if strDebug and comCode = "2" then
	response.Write("check if sessCode matches..<br>")
	end if
if comCode = "2" and sessCode = "67395ts02"  then
	if strDebug then
	response.Write("comCode = 2, sessCode matches..<br><br>")
	response.Write("check fsoObj and localhost vars..<br>")
	end if
	'instType = chkString(request.Form("installType"),"sqlstring")
    if fsoObj = true and localhost = 12 then
	  if strDebug then
	  'response.Write("fsoObj = true and localhost = 1. Lets check upload folders..<br>")
	  end if
	  'fsoMSG = chkUFolders()
	  'session.Contents("fsoMSG") = fsoMSG
	  if strDebug then
	  'response.Write("Upload folders checked:<br> " & fsoMSG & "<br><br>")
	  end if
    end if
	  if strDebug then
	  response.Write("Select install type...<br>")
	  end if
	select case instType
		case "new"
	  		if strDebug then
	  		response.Write("Install type: NEW...<br>")
	  		end if
			if request.Form("adminPass") <> request.Form("adminPass2") then
				session.Contents("erMsg") = "<font color=""#CC0000""><b>Your passwords didn't match</b></font>"
				showInstall()
			else
				createDB()
			end if
		case "upgrade13"
	  		if strDebug then
	  		response.Write("Install type: upgrade13...<br>")
	  		end if
			update13x()
			update20x21()
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
		case "upgrade15b3"
	  		if strDebug then
	  		response.Write("Install type: upgrade15b3...<br>")
	  		end if
			update15b3()
			update20x21()
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
		case "upgrade20"
	  		if strDebug then
	  		response.Write("Install type: upgrade20...<br>")
	  		end if
			update20x21() 
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
		case "upgrade21"
	  		if strDebug then
	  		response.Write("Install type: upgrade 21...<br>")
	  		end if
			update_211xRC1() 
			update_rc1_rc2()
			update_rc2_rc3()
		case "upgradeRC1"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC1...<br>")
	  		end if
			update_rc1_rc2()
			update_rc2_rc3()
		case "upgradeRC2"
	  		if strDebug then
	  		response.Write("Install type: upgrade RC2...<br>")
	  		end if
			update_rc2_rc3()
	end select
	  		if ErrorCount = 0 then
	  		  response.Write("<b>" & txtSUInstComp & "</b><br>")
			  Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
			else
	  		  response.Write("<font color=""#FF0000""><b>" & txtSUInstCompErr & "</b></font><br><br>")
	  		  response.Write("<br><b>ErrorCount: " & ErrorCount & "</b><br><br>")
	  		end if
			session.Contents("setup") = ""
	if ErrorCount = 0 then
	  'response.Redirect("site_setup.asp?cmd=3")
	  response.write("<br><br><a href=""site_setup.asp?cmd=3""><h4>" & txtSUContSetup & "</h4></a>")
	  response.Write "<br><br>" & dbHits & " " & txtSUDBHits & "<br><br>"
	else
	  'response.write("<br><br><a href=""site_setup.asp?cmd=3""><h4>" & txtSUContSetup & "</h4></a>")
	end if
else
	if ErrorCount = 0 then
		'if sInstallType <> "" then
		'response.Write("sInstallType: " & sInstallType)
		'response.Write("<br>v21Tbl: " & v21Tbl)
		'response.Write("<br>v20Tbl: " & v20Tbl)
		'response.Write("<br>oldTbl: " & oldTbl)
		'response.Write("<br>newTbl: " & newTbl)
		'response.Write("<br>betaTbl: " & betaTbl)
		'response.Write("<br>sInstallType: " & sInstallType)
		'response.Write("<br>sInstallType: " & sInstallType)
		if sInstallType = "" then
		  'has both tables
			Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
			showInstalled()
			session.Contents("setup") = ""
		else
			showInstall()
		end if 
	else
		errDisplay()
	end if
end if
%>
</td></tr>

            </table>
			</div>
          </div>
        </div>
        <div class="spThemeBlock1_bl"> 
          <div class="spThemeBlock1_br"> 
            <div class="spThemeBlock1_bc"></div>
          </div>
        </div>
      </div>


</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" align="center" width="100%">
<tr>
<td align="left" class="sp_FootLeft"></td>
<td align="left" class="sp_FootTile" nowrap><font face="Verdana, Arial, Helvetica" size="1"><a href="privacy.asp">Privacy</a></font></td>
<td align="right" class="sp_FootTile" nowrap><font face="Verdana, Arial, Helvetica" size="1">© 2005-2006 SkyPortal.net&nbsp;<%= txtSUAllRtsReserved %>.</font></td>
<td align="right" class="sp_FootTile" nowrap><font face="Verdana, Arial, Helvetica" size="1">
<a href="http://www.SkyPortal.net" title="Powered By: SkyPortal.net Version <%= strVer %>" target="_blank">
SkyPortal.net</a></font></td>
<td width="20" class="sp_FootTile" nowrap><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" height=15 width=15 border="0" align="right" alt="<%= txtSUGoTopPg %>"></a></td>
<td class="sp_FootRight"></td></tr>
</table>
</td>
</tr>
</table>
</td></tr></table></form>
</body>
</html>
<%
sub CheckSqlError(typ)
		'response.Write("<br>CheckSqlError: " & typ)
		'response.Write("<br>Errors.Count: " & my_Conn.Errors.Count)
  dim ChkConnErrorNumber
  if my_Conn.Errors.Count <> 0 or Err.number > 0 then  
	  for counter = 0 to my_Conn.Errors.Count '-1
		ChkConnErrorNumber = my_Conn.Errors(counter).Number
			my_Conn.Errors.Clear 
			Err.Clear
			session.Contents("setup") = "67395ts02"
		  	bHasTable = false
			select case typ
			  case "new"
				newTbl = 1
				sInstallType = typ
			  case "v13x"
				oldTbl = 1
				sInstallType = typ
			  case "v15b3"
				betaTbl = 1
				sInstallType = typ
			  case "v20"
				v20Tbl = 1
				sInstallType = typ
			  case "v21"
				v21Tbl = 1
				sInstallType = typ
			  case "vRC1"
				vRC1Tbl = 1
				sInstallType = typ
			  case "vRC2"
				vRC2Tbl = 1
				sInstallType = typ
			end select
		next
		  'bHasTable = false
		else
		  bHasTable = true
		end if
end sub

Sub showInstalled()
	 Err.Clear
 %>
<!--meta http-equiv="Refresh" content="3; URL=default.asp"--><br><br><br>
<table border="1" bgColor="#EAFFFF" cellspacing="0" cellpadding="5" width="80%" height="50%" align="center" bordercolor="#FFFFFF">
	<tr>
		<td align="center">
		<p>
		<font face="Verdana, Arial, Helvetica" size="3">
		<% If trim(fsoMSG) <> "" Then %>
		<ul><%= fsoMSG %></ul><br><br>
		<% End If %>
		<b><%= txtSUCongrats %></b><br><br>
		<%= txtSUSetupComp %><br><br></font>
	 <% 'chkInstallWrngs() %><br>
		<font face="Verdana, Arial, Helvetica" size="2">
		<b><%= txtSUChkGenSet %></b></font></p></td>
	</tr>
	<tr>
		<td align="center">
		<font face="Verdana, Arial, Helvetica" size="2">
		<a href="default.asp" target="_top"><%= txtSUContinue %>&nbsp;>>></a>
		</font></td>
	</tr>
</table><br><br>
<% end sub

sub shoModules() %>
  <p><b><%= txtSUUpgMods %></b><br>
  <%= txtSUUpdMods2 %></p>
  <fieldset style="width:200px;padding:5px;"><legend><%= txtSUMods %></legend>
  <table border="0" cellPadding="2" cellSpacing="0">
    <tr> 
      <td width="20%" align="right" vAlign="top"> 
        <input type="checkbox" name="Articles" value="1" checked>&nbsp;
      </td>
      <td width="80%"><%= txtArticles %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Classifieds" value="1" checked>&nbsp;
      </td>
      <td><%= txtClassifieds %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Downloads" value="1" checked>&nbsp;
      </td>
      <td><%= txtDownloads %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Forums" value="1" checked>&nbsp;
      </td>
      <td><%= txtForums %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Links" value="1" checked>&nbsp;
      </td>
      <td><%= txtLinks %></td>
    </tr>
    <tr>
      <td align="right" vAlign=top>
        <input type="checkbox" name="Pictures" value="1" checked>&nbsp;
      </td>
      <td><%= txtPics %></td>
    </tr>
  </table></fieldset>
<%
end sub

Sub showInstall() 
  if fsoObj and not localhost then
	'fsoMSG = chkUFolders()
  end if %>
  <table width="100%" border="1" cellPadding="8" cellSpacing="0" bordercolor="#FFFFFF">
	<tbody>
    <tr align="center"> 
      <td height=15 vAlign=top><font face="Verdana, Arial, Helvetica" size="2">
	    <% If session.Contents("erMsg") <> "" Then %>
		<br><%= session.Contents("erMsg") %><br><br>
		<% End If %>
		<% If sInstallType = "new" Then %>
			  <b><%= txtSUGetStart %></b><br><br>
        	  <font color="#0000FF"><b><%= txtSUNewInst %></b></font><br><br>
              <input type="hidden" name="installType" value="new"></font>
		<% Else %>
			  <h4><%= txtSUUpgradeTo %>&nbsp; v<%= strVer %>?</h4>
			  <% 
			  select case sInstallType
				'case "new"
				case "v13x"%>
                  <font color="#0000FF"><b>
				  <%= txtUpgrFrom %>&nbsp;MWP v1.3x</b></font>
                  <input type="hidden" name="installType" value="upgrade13"><%
				case "v15b3" %>
                  <font color="#0000FF"><b>
				  <%= txtUpgrFrom %>&nbsp;MWP.info v1.5 beta3</b></font>
                  <input type="hidden" name="installType" value="upgrade15b3"><%
				case "v20" %>
                  <font color="#0000FF"><b>
				  <%= txtUpgrFrom %>&nbsp;MWP.info v2.0</b></font>
                  <input type="hidden" name="installType" value="upgrade20"><%
				case "v21" %>
                  <font color="#0000FF"><b>
				  <%= txtUpgrFrom %>&nbsp;MWP.info v2.1</b></font>
                  <input type="hidden" name="installType" value="upgrade21">
				  <%
				case "vRC1" %>
                  <font color="#0000FF"><b>
				  <%= txtUpgrFrom %>&nbsp;SkyPortal vRC1</b></font>
                  <input type="hidden" name="installType" value="upgradeRC1">
				  <%
				case "vRC2" %>
                  <font color="#0000FF"><b>
				  <%= txtUpgrFrom %>&nbsp;SkyPortal vRC2</b></font>
                  <input type="hidden" name="installType" value="upgradeRC2">
				  <%
			  end select
			  %>
    <table width="100%" border="0" cellPadding=8 cellSpacing=0>
      <tr align="center"> 
      	<td height=15 vAlign=top></td>
	  </tr>
	</table>
		<% End If %>
	  </td></tr>
    <tr align="center"> 
      <td height=15 vAlign=top>
		<font face="Verdana, Arial, Helvetica" size="2"><b><%= txtSUPortRoot %>:&nbsp;&nbsp;
		<font color="#0000FF"><%= portalUrl %></font><br>
		<%= txtDBType %>:&nbsp;<font color="#FF0000"><%= strDBType %></font></b></font>
		<font face="Verdana, Arial, Helvetica" size="2">
		<% If fsoObj Then %>
		<br><br><b><%= txtSUFSOIsInst %></b>
		<% Else %>
		<br><br><b><%= txtSUFSOnotInst %></b>
		<% End If %>
		<% If fsoObj Then %>
			<br><ul><b><%= fsoMSG %></b></ul><br>
		<% End If %>
		</font>
      </td></tr>
	</tbody>
  </table>
<% If sInstallType <> "new" Then
     If sInstallType="v13x" or sInstallType="v15b3" or sInstallType="v20" or sInstallType="v21" Then
		shoModules()
	  end if %>
	<p><%= txtSUClkBtn %></p>
<% Else %>
	<div id="newInstall" style="display:block;">
    <table width="100%" border="1" cellPadding=8 cellSpacing=0 bordercolor="#FFFFFF">
      <tr align="center"> 
      	<td height=15 vAlign=top bgcolor="#F1F1F4">
		  <font size="2" face="Verdana, Arial, Helvetica"><b><%= txtSUFillOutFrm %></b></font></td></tr>
      <tr align="center">
        <td height=15 vAlign=top bgcolor="#F1F1F4">
		  <font size="2" face="Verdana, Arial, Helvetica"><%= txtSUSiteName %></font><br><br>
          <input class="textbox" name="siteName" type="text" id="siteName" value="<%= txtSUMySite %>">
        </td></tr>
      <tr> 
        <td height=15 align="center" vAlign=top bgcolor="#F1F1F4">
		<font face="Verdana, Arial, Helvetica" size=2>
		<%= txtSUSAName %>: <b><%= left(strWebMaster, instr(strWebMaster,",")-1) %></b>
		<br><br><%= txtSUSameNameAs %></font><br><br>
        <input class="textbox" name="adminName" type="hidden" id="adminName" value="<%= left(strWebMaster, instr(strWebMaster,",")-1) %>">
                                    </td>
                                </tr>
                                <tr> 
                                  <td width="31%" height=15 align="center" vAlign=top bgcolor="#F1F1F4">
								  <font face="Verdana, Arial, Helvetica" size=2>
								  <%= txtSUDefPass %><br>
                                    <br>
                                    <input class="textbox" name="adminPass" type="password" id="adminPass" value="<%= txtSUPassAdmin %>">
                                    <br>
                                    <br>
                                    <%= txtSUEnterPassAgin %><br>
                                    <br>
                                    <input class="textbox" name="adminPass2" type="password" id="adminPass2" value="<%= txtSUPassAdmin %>">
                                    </font></td>
                                </tr>
                                <tr>
                                  <td height=15 align="center" vAlign=top bgcolor="#F1F1F4">
								  <font face="Verdana, Arial, Helvetica" size=2>
								  <%= txtSUDetEmlComp %><br><br>
								  <%= txtSUSelEmlComp %><br><br>
                                    <% getEmailComponents() %>
                                    </font></td>
                                </tr>
                                <tr>
                                  <td height=15 align="center" vAlign=top bgcolor="#F1F1F4"><font face="Verdana, Arial, Helvetica" 
                        size=2><%= txtSUEmlServer %><br>
                                    <br>
                                    <input class="textbox" name="mailServer" type="text" value="<%= txtSUExEmlServer %>">
                                    </font></td>
                                </tr>
                                <tr> 
                                  <td height=15 align="center" vAlign=top bgcolor="#F1F1F4"><font face="Verdana, Arial, Helvetica" 
                        size=2><%= txtSUSiteEmlAdd %>: <br>
                                    <br>
                                    <input class="textbox" name="emailAddy" type="text" value="<%= txtSUExSiteEmlAdd %>">
       </font></td>
  	</tr></table></div>
<% End If %>
	<table width="100%" border="1" cellPadding=8 cellSpacing=0 bordercolor="#FFFFFF">
                              <!--<tr> 
                                <td height=15 align="center" vAlign=top bgcolor="#F1F1F4"> 
								<font face="Verdana, Arial, Helvetica" size=2>
								Is this being installed on<br>
								your home computer? <br>
                                    <br>
                                  <select name="localhost">
								  	<option value="0" Selected>Yes</option>
								  	<option value="1">No</option>
								  </select>
                                </td>
                              </tr>-->
      <tr> 
        <td height=15 align="center" vAlign=top> 
		  <input class="button" type="hidden" name="localhost" value="0">
          <input class="button" type="submit" name="Submit" value="<%= txtSUInstSkyPortal %>&nbsp;v<%= strVer %>">
        </td>
	  </tr>
	</table>
<% end sub

sub errDisplay()
%>
<table border="1" cellspacing="0" cellpadding="5" width="95%" align="center" bordercolor="#FFFFFF">
	<tr>
		<td bgColor=pink align="center">
		<font face="Verdana, Arial, Helvetica" size="4"><%= txtSUThereIsError %></font>
		<p>
		<font face="Verdana, Arial, Helvetica" size="2">
<%
Select case CustomCode
	 case 1 %>
		<%= txtSuCC1 %><br>	<br>
<% case 2 %>
		<%= txtSuCC2 %><br><br>
<% case 3 %>
		<%= txtSuCC3 %><br><br>
<% case 4 %>
		<%= txtSuCC4 %><br><br>
<% case else %>
		<%= txtSuCC5 %><br>
		<br>
<%
end select

		if ErrorCode <> "" then 
			Response.Write("</p><p>" & txtSUErrCode & " :  " & ErrorCode & " ")
			Response.Write("</p><p>" & strDBPath & "</p>")
		end if
%>
		</font></p></td>
	</tr>
	<tr>
		<td align="center"><font face="Verdana, Arial, Helvetica" size="2"><a href="site_setup.asp" target="_top"><%= txtSUClikToRetry %></a></font></td>
	</tr>
</table>
<%
End sub

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
				'Response.Write "NOT FOUND: ASP.NET Server Component<br>"
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
      objHttp.Send ""
		if err.number = 0 then
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
	   'Response.Write "FOUND: FileSystemObject scripting component<br>"
	   inFSOenabled = true
	   set fso = nothing
	 else 
	   'Response.Write "NOT FOUND: FileSystemObject scripting component<br>"
	   inFSOenabled = false
	 end if
     on error goto 0
	 fsoCheck = inFSOenabled
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

'Dim i
'for i=0 to UBound(arrComponent)
'	if isInstalled(arrComponent(i)) then
'	end if
'next

Response.Write("<select name=""emailComponent"">") & vbcrlf
'Response.Write("<ul>") & vbcrlf
'Response.Write("<option value=""none"" selected></option>") & vbcrlf
Dim i
for i=0 to UBound(arrComponent)
	if isInstalled(arrComponent(i)) then
	  'Response.Write("<li>"  & arrName(i) &"</li>") & vbcrlf
	  Response.Write("<option value=""" & arrValue(i) & """>" & arrName(i) &"</option>") & vbcrlf
	end if
next
'Response.Write("</ul>") & vbcrlf
Response.Write("</select>") & vbcrlf
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

function testFolder(fldr)
		dim fs
     	set fs = CreateObject("Scripting.FileSystemObject")
		set f = fs.GetFolder(fldr)
		'tmpMSG = tmpMSG & "<li></li>"
		fParent = right(f.ParentFolder,len(f.ParentFolder)-instrrev(f.ParentFolder,"\"))
		fs.CreateFolder fldr & "\test1"
		If fs.FolderExists(fldr & "\test1") = true Then
		  fs.CreateFolder fldr & "\test1\test2"
		  If fs.FolderExists(fldr & "\test\test2") = true Then
			'tmpMSG = tmpMSG & "<li>/files/" & x.Name & "/test folder created</li>"
			fs.DeleteFolder fldr & "\test\test2"
			If fs.FolderExists(fldr & "\test\test2") = true Then
				'tmpMSG = tmpMSG & "<li>/files/" & x.Name & "/test folder not deleted</li>"
				tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
				tmpMSG = tmpMSG & "Please check the <b>""" & fParent & "/" & f.Name & """</b> folder<br>for <b>""Delete""</b> permissions"
				tmpMSG = tmpMSG & "<br>" & fParent & "/" & f.Name & "/test/test2 folder was not deleted"
				tmpMSG = tmpMSG & "</span></li>"
				boolPerm = false
			else
				tmpMSG = tmpMSG & "<li><b>""" & fParent & "/" & f.Name & "/test""</b> - correctly set</li>"
				'boolPerm = false
			end if
			fs.DeleteFolder fldr & "\test"
			If fs.FolderExists(fldr & "\test") = true Then
				'tmpMSG = tmpMSG & "<li>/files/" & x.Name & "/test folder not deleted</li>"
				tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
				tmpMSG = tmpMSG & "Please check the <b>""" & fParent & "/" & f.Name & """</b> folder for <b>""Delete""</b> permissions"
				tmpMSG = tmpMSG & "<br>" & fParent & "/" & f.Name & "/test folder was not deleted"
				tmpMSG = tmpMSG & "</span></li>"
				boolPerm = false
			else
				tmpMSG = tmpMSG & "<li><b>""" & fParent & "/" & f.Name & """</b> - correctly set</li>"
				'boolPerm = false
			end if
		  else ':: \test\test2 not created
			'response.Write("test folder not created<br>")
			tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
			tmpMSG = tmpMSG & "Please check <b>""" & fParent & "/" & f.Name & """</b> folder permissions"
			tmpMSG = tmpMSG & "<br>Make sure that the permissions apply to child folders.</span></li>"
			boolPerm = false
		  end if
		else
			'response.Write("test folder not created<br>")
			tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
			tmpMSG = tmpMSG & "Please check <b>""" & fParent & "/" & f.Name & """</b> folder permissions"
			tmpMSG = tmpMSG & "<br>Make sure that the permissions apply to child folders.</span></li>"
			boolPerm = false
		end if
		set fs = nothing
end function

function chkDB(ckDBpath)
	tmpMSG = ""
	boolPerm = true
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(ckDBpath) = true Then
	  set fo = fso.GetFile(ckDBpath)
	  if fo.Name = "db2k30_x.mdb" then
		tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
		tmpMSG = tmpMSG & "Please <b>rename your Database</b> from the default of db2k30_x.mdb</span></li>"
	  else
		tmpMSG = tmpMSG & "<li>Database has been renamed!</li>"
	  end if
	  set fo = nothing
	else
		tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
		tmpMSG = tmpMSG & "Database does not exist, Check that<br>"
		tmpMSG = tmpMSG & "your Database path is correct:<br>"
		tmpMSG = tmpMSG & "<b>" & ckDBpath & "</b>"
		tmpMSG = tmpMSG & "</span></li>"
		boolPerm = false
	end if
	set fso = nothing
	chkDB = tmpMSG
end function

function chkPerm(ckFolder)
	tmpMSG = ""
	boolPerm = true
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(ckFolder) = true Then
	  set fo = fso.GetFolder(ckFolder) ':: get the "files" folder
	  testFolder(fo.Path)
	  for each x in fo.SubFolders
	    testFolder(x.Path)
	  next
	else
		tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
		tmpMSG = tmpMSG & "<b>""" & ckFolder & """</b> does not exist"
		tmpMSG = tmpMSG & "</span></li>"
		boolPerm = false
	end if
	if boolPerm = true then
	  tmpMSG = tmpMSG & "<li><b>""/" & fo.Name & """</b> folder permissions are correctly set</li>"
	end if
	set fo = nothing
	set fso = nothing
	chkPerm = tmpMSG
end function

function chkInstallWrngs()
det = DetectDotNetComponent("includes/scripts/checkfordotnet.aspx") %>
<hr><h4>
<% if det <> "" and inFSOenabled then %>
  SkyPortal can be fully 
  used on this server!<br><br>
<% else %>
	  SkyPortal's full features are not available.<br><br>
<%
   end if %>
</h4><%	  
	  if inFSOenabled = false then
	    response.Write("FileSystemObject is not available on this server<br>")
	    response.Write("Uploads will NOT be available in this installation")
	  else
	    response.Write("FileSystemObject is available on this server<br>")
	    response.Write("Uploads will be available in the installation<br><br>")
	  end if
      if det = "" and inFSOenabled = true then 
  		response.Write("ASP.NET is NOT installed on this server.<br>")
		response.Write("The image thumbnails will NOT be available.<br><br>")
	  elseif inFSOenabled = true then
  		response.Write("ASP.NET is installed on this server.<br>")
		response.Write("The image resizing will be available.<br><br>")
 	  end if
	  if inFSOenabled = true then
	    response.Write("<hr><div style=""text-align:left;padding-left:50px;"">")
	    response.Write("<h4>Installation Check:</h4>")
	    response.Write("<ol>")
		response.Write(chkDB(strDBPath))
		response.Write(chkPerm(Server.MapPath("/files")))
	    response.Write("</ol></div>")
 	  end if
end function

 %>
<!--#include file="install/createUpgrade.asp" -->
<!--#include file="install/create211_SP.asp" -->
<!--#include file="install/createCore.asp" -->