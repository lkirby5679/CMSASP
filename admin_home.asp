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
pgType = "SiteConfig"
configSys = true
%><!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/core_admin.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<% 'If "dogg" = "dogg" Then %>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<!--#INCLUDE file="includes/inc_admin_functions.asp" -->
<%
if Request.Form("Method_Type") = "configSys" then 
		Err_Msg = ""
		if Request.Form("strTitleImage") = "" then 
			Err_Msg = Err_Msg & txtMstEntrAddrssTimg
		end if
		if Request.Form("strHomeURL") = "" then 
			Err_Msg = Err_Msg & txtMstEntrURLHP
		end if
		if (left(lcase(Request.Form("strHomeURL")), 7) <> "http://" and left(lcase(Request.Form("strHomeURL")), 8) <> "https://") and Request.Form("strHomeURL") <> "" then
			Err_Msg = Err_Msg & txtMstPfxURLhttp
		end if
		if (right(lcase(Request.Form("strHomeURL")), 1) <> "/") then
			Err_Msg = Err_Msg & txtMstEndFURLbb
		end if
		if Request.Form("strEmailValx") < 1 then 
			Err_Msg = Err_Msg & txtMstChsNtfcnTpReg
		else
			stEmail = Request.Form("strEmailValx")
		end if
		if Request.Form("strAuthType") <> strAuthType and strAuthType = "db" then 
						
			if not(mLev = 4) then
				Err_Msg = Err_Msg & txtOnlyAdminChgAuth
			else
				call NTauthenticate()
				if Session(strCookieURL & "userid") = "" then
					Err_Msg = Err_Msg & txtEnblnAnonAccsSvrFst
				else	
					strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
					strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'; "

					executeThis(strSql)			
					call NTauthenticate()
					call NTUser()	
				end if
			end if
		end if
		if (Request.Form("strAuthType") <> strAuthType) and strAuthType = "nt" then 
			if not(mLev = 4) then
				Err_Msg = Err_Msg & txtOnlyAdminChgAuth
			else
				Session(strCookieURL & "Approval") = "" 
			end if	
		end if
		if Err_Msg = "" then
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRSITETITLE = '" & ChkString(Request.Form("strSiteTitle"),"sqlstring") & "'"
			strSql = strSql & ", C_STRCOPYRIGHT = '" & ChkString(Request.Form("strCopyright"),"sqlstring") & "'"
			strSql = strSql & ", C_STRTITLEIMAGE = '" & ChkString(Request.Form("strTitleImage"),"url") & "'"
			strSql = strSql & ", C_STRHOMEURL = '" & ChkString(Request.Form("strHomeURL"),"url") & "'"
			strSql = strSql & ", C_STRALLOWHTML = " & Request.Form("strAllowHTML") & ""
			strSql = strSql & ", C_STRALLOWFORUMCODE = " & Request.Form("strAllowForumCode") & ""
			strSql = strSql & ", C_STRICONS = " & Request.Form("strIcons") & ""
			strSql = strSql & ", C_STRSECUREADMIN = " & Request.Form("strSecureAdmin") & ""
			strSql = strSql & ", C_STRNOCOOKIES = " & Request.Form("strNoCookies") & ""
			strSql = strSql & ", C_STRFLOODCHECK = " & Request.Form("strFloodCheck") & ""
			strSql = strSql & ", C_STRFLOODCHECKTIME = " & Request.Form("strFloodCheckTime") & ""
			strSql = strSql & ", C_STRAUTHTYPE = '" & Request.Form("strAuthType") & "'"
			strSql = strSql & ", C_STRGFXBUTTONS = " & Request.Form("strGfxButtons")
			strSql = strSql & ", C_STRNEWREG = " & Request.Form("strNewReg")
			strSql = strSql & ", C_PMTYPE = " & Request.Form("strPMtype")
			strSql = strSql & ", C_STRLOCKDOWN = " & Request.Form("strLockDown")
			if intUploads = 1 then
			strSql = strSql & ", C_ALLOWUPLOADS = " & Request.Form("strAllowUploads")
			else
			strSql = strSql & ", C_ALLOWUPLOADS = 0"
			end if
			strSql = strSql & ", C_STRHEADERTYPE = " & Request.Form("strHeaderType")
			strSql = strSql & ", C_STRLOGINTYPE = " & Request.Form("strLoginType")
			strSql = strSql & ", C_STREMAILVAL = " & stEmail
			strSql = strSql & ", C_STRUNIQUEEMAIL = " & Request.Form("strUniqueEmail")
			strSql = strSql & ", C_SECIMAGE = " & Request.Form("strSecImage")			
			strSql = strSql & " WHERE CONFIG_ID = 1"
			response.Write(strsql & "<br><br>")
			executeThis(strSql)
			'response.End()
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""

			Session.Contents("adminHome") = txtMnConfigUpd
		else
			Err_Msg1 = txtProbDetails
			Session.Contents("adminHome") = Err_Msg1 & Err_Msg
		end if
    closeAndGo("admin_home.asp?cmd=1")
end if

if Request.Form("Method_Type") = "badWords" then 
		Err_Msg = ""
		if (Request.Form("strBadWordFilter") = "1" and strBadWordFilter = "1") or (Request.Form("strBadWordFilter") = "1" and strBadWordFilter = "0") then
			if Request.Form("strBadWords") = "" then 
				Err_Msg = Err_Msg & txtMstEntrWdsBdWdFltr
			end if
		end if

		if Err_Msg = "" then

			'
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRBADWORDFILTER = " & Request.Form("strBadWordFilter") & ""
			'if Request.Form("strBadWordFilter") = "1" then
				strSql = strSql & ", C_STRBADWORDS = '" & Request.Form("strBadWords") & "'"
			'end if
			strSql = strSql & " WHERE CONFIG_ID = " & 1
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
			Session.Contents("adminHome") = txtBdWdFltrUpdtd
		else 
			Err_Msg1 = txtProbDetails
			Session.Contents("adminHome") = Err_Msg1 & Err_Msg
		end if
    closeAndGo("admin_home.asp?cmd=2")
end if

if Request.Form("Method_Type") = "dateTime" then 
		Err_Msg = ""
		if Err_Msg = "" then
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRTIMETYPE              = '" & Request.Form("strTimeType") & "', "
			strSql = strSql & "     C_STRTIMEADJUST            = " & Request.Form("strTimeAdjust") & " "
			strSql = strSql & " WHERE CONFIG_ID = " & 1
		
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""

			Session.Contents("adminHome") = txtSvrDtTmCfgUpdtd
		else 
			Err_Msg1 = txtProbDetails
			Session.Contents("adminHome") = Err_Msg1 & Err_Msg
		end if
    closeAndGo("admin_home.asp?cmd=3")
end if

if Request.Form("Method_Type") = "emailServer" then 
		Err_Msg = ""
		if Request.Form("strMailServer") = "" and Request.Form("strEmail") = "1" and Request.Form("strMailMode") <> "cdonts" then 
			Err_Msg = Err_Msg & txtMstEntrAddrMailSvr
		end if
		if ((lcase(left(Request.Form("strMailServer"), 7)) = "http://") or (lcase(left(Request.Form("strMailServer"), 8)) = "https://") or Request.Form("strMailServer") = "") and Request.Form("strEmail") = "1" and Request.Form("strMailMode") <> "cdonts" then
			Err_Msg = Err_Msg & txtNoPrfxMailSvrHttp
		end if
		if Request.Form("strSender") = "" then 
			Err_Msg = Err_Msg & txtMstEntrEmlAddrAdmin
		else
			if EmailField(Request.Form("strSender")) = 0 and Request.Form("strSender") <> "" then 
				Err_Msg = Err_Msg & txtMstEntrVldEmailAddrAdmin
			end if
		end if
	
		if Err_Msg = "" then
			strSql = "UPDATE " & strTablePrefix & "CONFIG"
			strSql = strSql & " SET C_STREMAIL = " & Request.Form("strEmail") & ""
			strSql = strSql & ", C_STRMAILMODE = '" & Request.Form("strMailMode") & "'"
			'if Request.Form("strMailServer") <> "" then
				strSql = strSql & ", C_STRMAILSERVER = '" & Request.Form("strMailServer") & "'"
			'end if
			'if Request.Form("strSender") <> "" then
				strSql = strSql & ", C_STRSENDER = '" & Request.Form("strSender") & "'"
			'end if
			strSql = strSql & ", C_STRLOGONFORMAIL = " & Request.Form("strLogonForMail") & ""
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
			
			Session.Contents("adminHome") = txtEmlSvrCfgUpdtd
		else 
			Err_Msg1 = txtProbDetails
			Session.Contents("adminHome") = Err_Msg1 & Err_Msg
		end if
    closeAndGo("admin_home.asp?cmd=4")
end if

if Request.Form("Method_Type") = "ntConfig" then 
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRNTGROUPS = " & Request.Form("strNTGroups")
			strSql = strSql & ", C_STRAUTOLOGON = " & Request.Form("strAutoLogon")
			strSql = strSql & " WHERE CONFIG_ID = " & 1
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""

			Session.Contents("adminHome") = txtNTFeatCfgUpdtd
    closeAndGo("admin_home.asp?cmd=9")
end if %>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	menu_admin()
	spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
	  <% 
	  intSkin = getSkin(intSubSkin,2)
	  'breadcrumb here
  	  arg1 = txtAdminHome & "|admin_home.asp"
  	  arg2 = ""
  	  arg3 = ""
  	  arg4 = ""
  	  arg5 = ""
  	  arg6 = ""
  
  	  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	  spThemeBlock1_open(intSkin)
	  response.Write("<div id=""zz"" style=""display:block;"">")
	if Session.Contents("adminHome") <> "" then
	  response.Write("<p align=""center""><ul>")
	  response.Write(Session.Contents("adminHome"))
	  response.Write("</ul></p>")
	  Session.Contents("adminHome") = ""
	end if
	  response.Write("</div>")
		adminHome()
		generalConfig()
		badWords()
		dateTime()
		emailConfig()
		checkInstall()
		emailMembers()
		serverVar()
		siteVar()
		NTconfig() %>
	<% spThemeBlock1_close(intSkin) %>
    </td>
  </tr>
</table>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF

sub adminHome() %>
	<div id="aa" style="display:<%= aa %>;">
    <table border="0" width="100%" cellspacing="1" cellpadding="4" class="grid">
      <tr>
        <td class="tSubTitle" colspan=2><span class="fAltSubTitle"><b><% =txtAdminstFx %></b></span></td>
      </tr>
      <tr>
        <td class="tCellAlt1" valign=top>
        <p><b><% =txtPndItms %></b>
        <span class="fAlert"><ul>
<% if chkApp("events","USERS") then %>
		<% cntEvents = getCount("EVENT_ID",strTablePrefix & "EVENTS","PENDING=1")
		   If cntEvents <> 0 then %>
			    <li><b><a href="admin_events.asp"><%= cntEvents %>&nbsp;<%= txtNwEvtsApprv %></a></b></li>
		<% End IF %>
<% End IF %>
<% if chkApp("article","USERS") then %>
        <% cntArticle = getCount("ARTICLE_ID","ARTICLE","SHOW=0")
		   If cntArticle <> 0 then %>
			  <li><b><a href="admin_article_main.asp"><%= cntArticle %>&nbsp;<%= txtNwArtclsApprv %></a></b></li>
		<% End IF %>
<% End IF %>
<% if chkApp("classifieds","USERS") then %>
		<% cntClass = getCount("title","CLASSIFIED","SHOW=0")
			cntCbad = getCount("title","CLASSIFIED","BADLINK <> 0")
			If cntClass <> 0 or cntCbad <> 0 then %>
			  <b>
			  <% If cntClass <> 0 then %>
			  <li><a href="admin_classified_main.asp"><%= cntClass %>&nbsp;<%= txtNwClassifApprv %></a></li>
			  <% End IF %>
			  <% If cntCbad <> 0 then %>
              <li><a href="admin_classified_admin.asp?cmd=40"><%= cntCbad %>&nbsp;<%= txtBdClassif %></a></li>
			  <% End IF %>
			  </b>
		<% End IF %>
<% End IF %>
<% if chkApp("downloads","USERS") then %>
        <% cntDL = getCount("url","DL","SHOW=0")
		   cntDLB = getCount("url","DL","BADLINK <> 0")
		   If cntDL <> 0 or cntDLB <> 0 then %>
				<b>
				<% If cntDL <> 0 then %>
				<li><a href="admin_dl_main.asp"><%= cntDL %>&nbsp;<%= txtNwDLsApprv %></a></li>
				<% End IF %>
				<% If cntDLB <> 0 then %>
                <li><a href="admin_dl_admin.asp?cmd=40"><%= cntDLB %>&nbsp;<%= txtBdDLs %></a></li>
				<% End IF %>
				</b>
		<% End IF %>
<% End IF %>
<% if chkApp("links","USERS") then %>
        <% cntLink = getCount("url","LINKS","SHOW=0")
		   cntBad = getCount("url","LINKS","BADLINK <> 0")
		   If cntLink <> 0 or cntBad <> 0 then %>
				<b>
				<% If cntLink <> 0 then %>
				<li><a href="admin_links_main.asp"><%= cntLink %>&nbsp;<%= txtNwLinksApprv %></a></li>
				<% End IF %>
				<% If cntBad <> 0 then %>
                <li><a href="admin_links_admin.asp?cmd=40"><%= cntBad %>&nbsp;<%= txtBdLinks %></a></li>
				<% End IF %>
				</b>
		<% End IF %>
<% End IF %>
<% if chkApp("pictures","USERS") then %>
        <% cntPic = getCount("pic_ID","pic","SHOW=0")
		   cntPicB = getCount("pic_ID","pic","BADLINK <> 0")
		   If cntPic <> 0 or cntPicB <> 0 then %>
				<b>
				<% If cntPic <> 0 then %>
				<li><a href="admin_pic_main.asp"><%= cntPic %>&nbsp;<%= txtNwPicApprv %></a></li>
				<% End IF %>
				<% If cntPicB <> 0 then %>
                <li><a href="admin_pic_admin.asp?cmd=40"><%= cntPicB %>&nbsp;<%= txtBdPicts %></a></li>
				<% End IF %>
				</b>
		<% End IF %>
<% End IF %>
		</ul></span></p>
        </td>
        <td class="tCellAlt1" valign=top>
        <p><b><% =txtExtraSpace %></b></p>
        </td>
      </tr>
    </table>
	</div>
<%
end sub

sub generalConfig() %>
	<div id="ab" style="display:<%= ab %>;">
<% 
select case strEmailVal
  case 1
  	em1 = 1
	em2 = 1
  case 2
  	em1 = 1
	em2 = 2
  case 3
  	em1 = 1
	em2 = 3
  case 4
  	em1 = 1
	em2 = 4
  case 5
  	em1 = 2
	em2 = 2
  case 6
  	em1 = 2
	em2 = 4
  case 7
  	em1 = 3
	em2 = 2
  case 8
  	em1 = 4
	em2 = 2
end select
 %>
<form action="admin_home.asp" method="post" id="myChoices" name="myChoices">
<input type="hidden" name="Method_Type" value="configSys">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt2">
        <table border="0" cellspacing="1" cellpadding="1">
          <tr valign="top"> 
            <td class="tTitle" colspan="2"><b>&nbsp;<% =txtSiteConfig %></b></td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtSiteTitle %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="strSiteTitle" size="25" value="<% if strSiteTitle <> "" then Response.Write(strSiteTitle) else '## do nothing %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#forumtitle')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtCopyright %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="strCopyright" size="25" value="<% if strCopyright <> "" then Response.Write(strCopyright) else Response.Write("&copy;" & strSiteTitle) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#copyright')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
              <input type="hidden" class="textbox" name="strTitleImage" size="25" value="<% if strTitleImage <> "" then Response.Write(strTitleImage) else Response.Write("images/site_logo.jpg") %>">
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtHomeURL %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="strHomeURL" size="25" value="<% if strHomeURL <> "" then Response.Write(strHomeURL) else '## Do Nothing %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#homeurl')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtVersionInfo %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <% Response.Write "[<i>SkyPortal " & strWebSiteVersion & "</i>]" %>
               
              <!--<a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#forumtitle')"><img src="icon_smile_question.gif" border="0"></a>-->
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtAuthType %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <% =txtDB %>
              <input type="radio" class="radio" name="strAuthType" value="db" <% if strAuthType = "db" then Response.Write("checked") %>>
              NT: 
              <input type="radio" class="radio" name="strAuthType" value="nt" <% if strAuthType = "nt" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#AuthType')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
              </td>
          </tr>
  <tr valign="middle">
    <td class="tCellAlt0" align="right"><b><% =txtAllowHTML %></b> </td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strAllowHTML" value="1"<% if strAllowHTML <> "0" then Response.Write(" checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strAllowHTML" value="0"<% if strAllowHTML = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#AllowHTML')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="middle">
    <td class="tCellAlt0" align="right"><b><% =txtAllwFrmCode %></b> </td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strAllowForumCode" value="1"<% if strAllowForumCode <> "0" then Response.Write(" checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strAllowForumCode" value="0"<% if strAllowForumCode = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#AllowForumCode')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="middle">
    <td class="tCellAlt0" align="right"><b><% =txtIcons %>:</b> </td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strIcons" value="1" <% if (lcase(strIcons) <> "0" or lcase(Smiles) <> "0") then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strIcons" value="0" <% if (lcase(strIcons) = "0" or lcase(Smiles) = "0") then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#icons')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="middle">
    <td class="tCellAlt0" align="right"><b><% =txtSecrAdminMode %></b> </td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strSecureAdmin" value="1" <% if lcase(strSecureAdmin) <> "0" then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strSecureAdmin" value="0" <% if lcase(strSecureAdmin) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#secureadminmode')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="middle">
    <td class="tCellAlt0" align="right"><b><% =txtNonCookMode %></b> </td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strNoCookies" value="1" <% if lcase(strNoCookies) <> "0" then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strNoCookies" value="0" <% if lcase(strNoCookies) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#allownoncookies')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="middle">
    <td class="tCellAlt0" align="right"><b><% =txtFloodCheck %></b> </td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strFloodCheck" value="1"<% if strFloodCheck <> "0" then Response.Write(" checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strFloodCheck" value="0"<% if strFloodCheck = "0" then Response.Write(" checked") %>>
    <select name="strFloodCheckTime">
      <option<% if (lcase(strFloodCheckTime)="-10") then Response.Write(" selected") %> value="-10">10 <% =txtSeconds %></option>
      <option<% if (lcase(strFloodCheckTime)="-30") then Response.Write(" selected") %> value="-30">30 <% =txtSeconds %></option>
      <option<% if (lcase(strFloodCheckTime)="-60") then Response.Write(" selected") %> value="-60">60 <% =txtSeconds %></option>
      <option<% if (lcase(strFloodCheckTime)="-90") then Response.Write(" selected") %> value="-90">90 <% =txtSeconds %></option>
      <option<% if (lcase(strFloodCheckTime)="-120") then Response.Write(" selected") %> value="-120">120 <% =txtSeconds %></option>
    </select>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#FloodCheck')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><% =txtGfxButtons %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <% =txtOn %> 
              <input type="radio" class="radio" name="strGfxButtons" value="1" <% if strGfxButtons = "1" then Response.Write("checked") %>>
              <% =txtOff %> 
              <input type="radio" class="radio" name="strGfxButtons" value="0" <% if strGfxButtons = "0" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#GfxButtons')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtPvtMsgType %>&nbsp; </b></td>
            <td class="tCellAlt0"> 
              <select name="strPMtype">
                <option value="0"<% if strPMtype = "0" then Response.Write(" selected") %>> 
                <% =txtGraphic %> </option>
                <option value="1"<% if strPMtype = "1" then Response.Write(" selected") %>> 
                <% =txtToast %> </option>
                <option value="2"<% if strPMtype = "2" then Response.Write(" selected") %>> 
                <% =txtBoth %> </option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#PMtype')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><% =txtAllowUplds %></b>&nbsp;</td>
            <td class="tCellAlt0">
  			<select name="strAllowUploads"<% if intUploads = 0 then %> disabled="disabled"<% end if %>>
    			<option value="1"<% if strAllowUploads = 1 then Response.Write(" selected") %>><% =txtYes %></option>
    			<option value="0"<% if strAllowUploads = 0 then Response.Write(" selected") %>><% =txtNo %></option>
  			</select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#allowuploads')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtHeaderType %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <select name="strHeaderType">
                <option value=0<% If strHeaderType = 0 Then response.Write " selected" else 'do nothing %>><% =txtNone %></option>
                <option value=3<% If strHeaderType = 3 Then response.Write " selected" else 'do nothing %>><% =txtIcons %></option>
                <option value=1<% If strHeaderType = 1 Then response.Write " selected" else 'do nothing %>><% =txtRotatBanner %></option>
                <option value=2<% If strHeaderType = 2 Then response.Write " selected" else 'do nothing %>><% =txtRndBanner %></option>
                <option value=4<% If strHeaderType = 4 Then response.Write " selected" else 'do nothing %>><% =txtOther %></option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#headtype')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="middle"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtSiteLockDn %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <select name="strLockDown">
                <option value=0<% If strLockDown = 0 Then response.Write " selected" else 'do nothing %>><% =txtNo %></option>
                <option value=1<% If strLockDown = 1 Then response.Write " selected" else 'do nothing %>><% =txtYes %></option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#lockdown')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tTitle" colspan="2"><b>&nbsp;<% =txtRegist %>:</b></td>
          </tr>
          <tr valign="top"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtLoginBoxLctn %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <select name="strLoginType">
                <option value=0<% If strLoginType = 0 Then response.Write " selected" else 'do nothing %>><% =txtHeader %></option>
                <option value=1<% If strLoginType = 1 Then response.Write " selected" else 'do nothing %>><% =txtNavBar %></option>
                <option value=2<% If strLoginType = 2 Then response.Write " selected" else 'do nothing %>><% =txtOther %></option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#loginloc')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtNewRegs %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <% =txtUsers %> 
              <input type="radio" class="radio" name="strNewReg" id="on" value="1" <% if strNewReg = "1" then Response.Write("checked") %>>
              <% =txtAdmin %> 
              <input type="radio" class="radio" name="strNewReg" id="off" value="0" <% if strNewReg = "0" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#allowreg')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
              </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" align="right"><b><% =txtReqUnqEmail %></b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <% =txtOn %> 
              <input type="radio" class="radio" name="strUniqueEmail" value="1" <% if strUniqueEmail = "1" then Response.Write("checked") %>>
              <% =txtOff %> 
              <input type="radio" class="radio" name="strUniqueEmail" value="0" <% if strUniqueEmail <> "1" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#UniqueEmail')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
              </td>
          </tr>
          <tr valign="top"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtValidation %></b>&nbsp; 
            </td>
            <td class="tCellAlt0"> 
              <select id="strEmailVal" name="strEmailVal" onChange="selectChange(this, myChoices.strEmailValx, arrItems1, arrItemsGrp1);">
                <option value=1<% If em1= 1 Then response.Write(" selected") %>><% =txtNone %></option>
                <option value=2<% If em1= 2 Then response.Write(" selected") %>><% =txtMember %></option>
                <option value=3<% If em1= 3 Then response.Write(" selected") %>><% =txtAdmin %></option>
                <option value=4<% If em1= 4 Then response.Write(" selected") %>><% =txtMember %> & <% =txtAdmin %></option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#valtype')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="top"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtNotifs %></b>&nbsp; 
            </td>
            <td class="tCellAlt0"> 
              <select id="strEmailValx" name="strEmailValx">
                <!-- <option value=0>[SELECT]</option> -->
				<% If strEmailVal= 1 or strEmailVal= 2 or strEmailVal= 3 or strEmailVal= 4 Then %>
                <option value=1<% If em2= 1 Then response.Write(" selected") %>><% =txtNone %></option>
                <option value=2<% If em2= 2 Then response.Write(" selected") %>><% =txtMember %></option>
                <option value=3<% If em2= 3 Then response.Write(" selected") %>><% =txtAdmin %></option>
                <option value=4<% If em2= 4 Then response.Write(" selected") %>><% =txtMember %> & <% =txtAdmin %></option>
				<% ElseIf strEmailVal= 5 or strEmailVal = 6 Then %>
                <option value=5<% If em2= 2 Then response.Write(" selected") %>><% =txtMember %></option>
                <option value=6<% If em2= 4 Then response.Write(" selected") %>><% =txtMember %> & <% =txtAdmin %></option>
				<% Elseif strEmailVal= 7 Then %>
                <option value=7 selected><% =txtMember %></option>
				<% Else %>
                <option value=8 selected><% =txtMember %></option>				
				<% End If %>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#valtype')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
           <tr valign="top"> 
            <td class="tTitle" colspan="2"><b>&nbsp;<% =txtSecImgPrtctn %></b></td>
          </tr>
          <tr valign="top"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><% =txtSecImg %></b>&nbsp; 
            </td>
            <td class="tCellAlt0"> 
              <select id="strSecImage" name="strSecImage">
                <!-- <option value=0>[SELECT]</option> -->
                <option value=0<% If SecImage= 0 Then response.Write(" selected") %>><% =txtOff %></option>
                <option value=1<% If SecImage= 1 Then response.Write(" selected") %>><% =txtRegist %></option>
                <option value=2<% If SecImage= 2 Then response.Write(" selected") %>><% =txtUsers %></option>
                <option value=3<% If SecImage= 3 Then response.Write(" selected") %>><% =txtUsers %> & <% =txtAdmin %></option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#sectype')"><img src="images/icons/icon_smile_question.gif" border="0"></a> 
            </td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" colspan="2" align="center"> 
              <input type="submit" value="<% =txtSubmitNwCfg %>" id="submit1" name="submit1" class="button" title="<% =txtSubmitNwCfg %>"> <input type="reset" value="<% =txtResetOldVal %>" id="reset1" name="reset1" class="button" title="<% =txtResetOldVal %>">
            </td>
          </tr>
        </table>
    </td>
  </tr>
</table>
</form>
	</div>
<%
end sub

sub badWords() %>
	<div id="ac" style="display:<%= ac %>;">
<form action="admin_home.asp" method="post" id="FormBW" name="FormBW">
<input type="hidden" name="Method_Type" value="badWords">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b><% =txtBadWdCfg %></b></td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtBWfilter %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strBadWordFilter" value="1" <% if strBadWordFilter <> "0" then Response.Write("checked")%>> 
    <% =txtOff %> <input type="radio" class="radio" name="strBadWordFilter" value="0" <% if strBadWordFilter = "0" then Response.Write("checked")%>>
    <input type="text" name="strBadWords" size="20" value="<% if strBadWords <> "" then Response.Write(strBadWords) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#badwordfilter')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
   </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<% =txtSubmitNwCfg %>" id="submit1" name="submit1" class="button" title="<% =txtSubmitNwCfg %>"> <input type="reset" value="<% =txtResetOldVal %>" id="reset1" name="reset1" class="button" title="<% =txtResetOldVal %>"></td>

  </tr>
</table>
    </td>
  </tr>
</table></form>
	</div>
<%
end sub

sub dateTime()
	strSql = "SELECT C_STRTIMETYPE, C_STRTIMEADJUST FROM " & strTablePrefix & "CONFIG"
	strSql = strSql & " WHERE CONFIG_ID = 1"
	set rsTD = my_Conn.execute(strSql)
	  strTimeType = rsTD("C_STRTIMETYPE")
	  strTimeAdjust = rsTD("C_STRTIMEADJUST")
	  session.LCID = intPortalLCID
	  strCurDateAdjust = DateAdd("h", strTimeAdjust , Now()) 'portal offset from server
	  strCurDateString = DateToStr2(strCurDateAdjust)
	  strForumTimeAdjust = strToDate(strCurDateString)
	  strForumDateAdjust = ChkDate(strCurDateString)
	  strServerDateTime = strToDate(DateToStr2(Now()))
	set rsTD = nothing %>
	<div id="ad" style="display:<%= ad %>;">
<form action="admin_home.asp" method="post" id="formEle" name="Form1">
<input type="hidden" name="Method_Type" value="dateTime">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b><% =txtSvrDtTmCfg %></b></td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b>Server Time:</b>&nbsp;</td>
    <td class="tCellAlt0">
	<%= strServerDateTime %>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b>Portal Time:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <%= strForumTimeAdjust %>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b>Portal LCID:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <%= intPortalLCID %>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtTimeDsp %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <% =txt24hr %> <input type="radio" class="radio" name="strTimeType" value="24" <% if strTimeType = "24" then Response.Write("checked") %>> 
    <% =txt12hr %> <input type="radio" class="radio" name="strTimeType" value="12" <% if strTimeType = "12" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#timetype')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtTimeAdj %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <select name="strTimeAdjust">
	  <% 
	  for idt = -24 to 24 %>
      <option Value="<%= idt %>"<%= chkSelect(strTimeAdjust,idt) %>><%= idt %></option>
	  <% 
	  next %>
    </select>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#TimeAdjust')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td height="30" class="tCellAlt0" colspan="2" align="center" valign="middle"><input type="submit" value="<% =txtSubmitNwCfg %>" id="submit1" name="submit1" class="button" title="<% =txtSubmitNwCfg %>"> <input type="reset" value="<% =txtResetOldVal %>" id="reset1" name="reset1" class="button" title="<% =txtResetOldVal %>"></td>
  </tr>
</table>
    </td>
  </tr>
</table></form>
	</div>
<%
end sub

sub emailConfig() %>
	<div id="ae" style="display:<%= ae %>;">
<form action="admin_home.asp" method="post" id="formEle" name="Form1">
<input type="hidden" name="Method_Type" value="emailServer">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b><% =txtEmlSvrCfg %></b></td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtEmailMode %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strEmail" value="1" <% if lcase(strEmail) <> "0" then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strEmail" value="0" <% if lcase(strEmail) = "0" then Response.Write("checked") %>>
	<% getEmailComponents() %>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#email')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtEmailSvrAddr %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <input type="text" name="strMailServer" size="25" value="<% =strMailServer %>">
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#mailserver')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtAdminEmailAddr %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <input type="text" name="strSender" size="25" value="<% =strSender %>">
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#sender')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtReqLogonEml %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strLogonForMail" value="1" <% if strLogonForMail = "1" then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strLogonForMail" value="0" <% if strLogonForMail <> "1" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#LogonForMail')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<% =txtSubmitNwCfg %>" id="submit1" name="submit1" class="button" title="<% =txtSubmitNwCfg %>"> <input type="reset" value="<% =txtResetOldVal %>" id="reset1" name="reset1" class="button" title="<% =txtResetOldVal %>"></td>
  </tr>
</table>
    </td>
  </tr>
</table></form>
	</div>
<%
end sub

sub checkInstall() %>
	<div id="af" style="display:<%= af %>;">
	</div>
<%
end sub

sub emailMembers() %>
	<div id="ag" style="display:<%= ag %>;">
	</div>
<%
end sub

sub serverVar() %>
	<div id="ah" style="display:<%= ah %>;">
	<p><br><% =txtSvrInfoTxt %></p>
    <table border="0" cellspacing="1" cellpadding="1" align=center width="95%">
      <tr>
        <td class="tSubTitle"><b><% =txtVarName %></b></td>
        <td class="tSubTitle"><b><% =txtValue %></b></td>
      </tr>
<% for each key in Request.ServerVariables %>
      <tr>
        <td class="tCellAlt0" valign="top"><b><% =key %></b></td>
        <td class="tCellAlt0"><%if Request.ServerVariables(key) = "" then
Response.Write "&nbsp;"
else
Response.Write Request.Servervariables(key)
end if 
%>
        </td>
      </tr>
<% next %>
    </table>
	</div>
<%
end sub

sub siteVar() %>
	<div id="ai" style="display:<%= ai %>;">
	<!--div id="ai" style="display:block;"-->

    <table border="0" cellspacing="1" cellpadding="1" align=center width="95%">
     <tr>
        <td class="tTitle"><b><% =txtVarName %></b></td>
        <td class="tTitle"><b><% =txtValue %></b></td>
      </tr>

     <tr>
        <td class="tSubTitle"  align="center" colspan="2" ><b><% =txtGralInfo %></b></td>
     </tr>	
     <tr>
        <td class="tCellAlt0"><b>strWebMaster</b></td>
        <td class="tCellAlt0"><%=ChkString(strWebmaster, "display")%></td>
      </tr>	
     <tr>
        <td class="tCellAlt0"><b>StrCookieUrl</b></td>
        <td class="tCellAlt0"><%=ChkString(StrCookieUrl, "display")%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>StrUniqueID</b></td>
        <td class="tCellAlt0"><%=ChkString(StrUniqueID, "display")%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>strAuthType</b></td>
        <td class="tCellAlt0"><%=ChkString(strAuthType, "display")%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>strDBNTSQLName</b></td>
        <td class="tCellAlt0"><%=ChkString(strDBNTSQLName, "display")%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>STRdbntUserName</b></td>
        <td class="tCellAlt0"><%=ChkString(STRdbntUserName, "display")%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>strDBType</b></td>
        <td class="tCellAlt0"><%=ChkString(strDBType, "display")%></td>
      </tr>	
     <tr>
        <td class="tCellAlt0"><b>strConnString</b></td>
        <td class="tCellAlt0"><%=ChkString(strConnString, "display")%></td>
      </tr>	
     <tr>
        <td class="tCellAlt0"><b>strTheme</b></td>
        <td class="tCellAlt0"><%=strTheme%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>FSOenabled</b></td>
        <td class="tCellAlt0"><%=FSOenabled%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>varBrowser</b></td>
        <td class="tCellAlt0"><%=varBrowser%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>pageTimer</b></td>
        <td class="tCellAlt0"><%=ChkString(pageTimer, "display")%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>strForumTimeAdjust</b></td>
        <td class="tCellAlt0"><%=strForumTimeAdjust%></td>
      </tr>
     <tr>
        <td class="tCellAlt0"><b>strForumDateAdjust</b></td>
        <td class="tCellAlt0"><%=strForumDateAdjust%></td>
      </tr>
     <tr>
        <td class="tSubTitle"  align="center" colspan="2" ><b><% =txtCookies %></b></td>
      </tr>
<% for each key in Request.Cookies 

	if Request.Cookies(key).HasKeys then
		for each subkey in Request.Cookies(key)
%>
 		     <tr>
		        <td class="tCellAlt0" valign="top"><b><% =ChkString(key, "display") %> (<% =ChkString(subkey, "display") %>)</b></td>
		        <td class="tCellAlt0">
<%
		if Request.Cookies(key)(subkey) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write CStr(Request.Cookies(key)(subkey))
		end if 
%>
		        </td>
		      </tr>
<%		next
	else
%>
 		     <tr>
		        <td class="tCellAlt0" valign="top"><b><% =ChkString(key, "display") %></b></td>
		        <td class="tCellAlt0">
<%
		if Request.Cookies(key) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write ChkString(CStr(Request.Cookies(key)), "display")
		end if 
%>
		        </td>
		      </tr>
<%
	end if
next  %>
 
     <tr>
        <td class="tSubTitle"  align="center" colspan="2"><b><% =txtSessVars %></b></td>
      </tr>
<% for each key in Session.Contents

	if left(key, len(strCookieUrl)) = strCookieUrl or left(key, len(strUniqueID)) = strUniqueID then
%>
      <tr>
        <td class="tCellAlt0" valign="top"><b><% =ChkString(key, "display") %></b></td>
        <td class="tCellAlt0">
<%
	if Session.Contents(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write ChkString(CStr(Session.Contents(key)), "display")
	end if 
%>
        </td>
      </tr>
<% 
	end if
next 

%>

      <tr>
        <td class="tSubTitle"  align="center" colspan="2" ><b><% =txtAppVars %></b></td>
      </tr>
<% for each key in Application.Contents
	  'StrIPGateWarnMsg = Application(strCookieURL & strUniqueID & "STRIPGATEWARNMSG")

	if left(key, len(strCookieUrl & strUniqueID)) = strCookieUrl & strUniqueID then
%>
      <tr>
        <td class="tCellAlt0" valign="top"><b><% = ChkString(key, "display") %></b></td>
        <td class="tCellAlt0">
<%
	if len(Application.Contents(key)) = 0 then
		Response.Write "&nbsp;"
	else
		Response.Write Application.Contents(key)
	end if 
%>
        </td>
      </tr>
<% 
	end if
next 

%>
    </table>
	</div>
<%
end sub

sub getEmailComponents()
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

Response.Write("<select name=""strMailMode"">") & vbcrlf
Dim ix
for ix=0 to UBound(arrComponent)
	if isInstalled(arrComponent(ix)) then
	  Response.Write("<option value=""" & arrValue(ix) & """")
		if lcase(strMailMode)=arrValue(ix) then
		  Response.Write(" selected")
		end if
	  Response.Write(">" & arrName(ix) & "</option>") & vbcrlf
	end if
next
Response.Write("</select>") & vbcrlf
end sub				'

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

sub NTconfig() %>
	<div id="aj" style="display:<%= aj %>;">
<form action="admin_home.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="ntConfig">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b><% =txtFeatNTcfg %></b></td>
  </tr>
<% if strAuthType = "nt" then %>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtUseNTgrps %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strNTGroups" value="1" <% if strNTGroups = "1" then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strNTGroups" value="0" <% if strNTGroups = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" align="right"><b><% =txtUseNTautoLgn %>:</b>&nbsp;</td>
    <td class="tCellAlt0">
    <% =txtOn %> <input type="radio" class="radio" name="strAutoLogon" value="1" <% if strAutoLogon = "1" then Response.Write("checked") %>> 
    <% =txtOff %> <input type="radio" class="radio" name="strAutoLogon" value="0" <% if strAutoLogon = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<% =txtSubmitNwCfg %>" id="submit1" name="submit1" class="button" title="<% =txtSubmitNwCfg %>"> <input type="reset" value="<% =txtResetOldVal %>" id="reset1" name="reset1" class="button" title="<% =txtResetOldVal %>"></td>
  </tr>
<% else %>
  <tr valign="top">
    <td class="fTitle" colspan="2"><p><b><% =txtMstHvNTauthOn %></b></p></td>
  </tr>
<% end if %>
</table>
    </td>
  </tr>
</table>
</form>
	</div>
<%
end sub

 %>