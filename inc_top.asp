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

const strWebSiteMVersion = "RC3"
tempArr = split(strWebMaster, ",")
strSiteOwner = tempArr(0)

dim strMBirthday, pmimage, pmCount

if not isObject(my_Conn) then
  on error resume next
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear
	my_Conn.Open strConnString
	'::Lets check to see if the strConnString or db path has changed
	if my_conn.Errors.Count <> 0 then 
		'we can't connect, lets display the error
		my_conn.Errors.Clear 
		set my_Conn = nothing
		Response.Redirect "site_setup.asp?RC=1"
	end if
  on error goto 0
end if
%>
<!--#include file="includes/inc_emails.asp" -->
<!--#include file="includes/inc_theme.asp" -->
<!--#include file="includes/inc_encryption.asp" -->
<!--#INCLUDE file="includes/fp_core.asp" -->
<!--#INCLUDE file="includes/default_menu.asp" -->
<%
showDaPage = true
  sScript = request.ServerVariables("SCRIPT_NAME")
if not uploadPg = true then
  strReferrer = chkString(Request.ServerVariables("HTTP_REFERER"), "refer")
  strReferrer = Replace(strReferrer, ">", "&gt;")
  strReferrer = Replace(strReferrer, "<", "&lt;")
  select case Request.Form("Method_Type")
	case "login"
		strDBNTFUserName = chkString(Request.Form("Name"),"display")
		mLev = getMlev(strDBNTFUserName, pEncrypt(pEnPrefix & Request.Form("Password")))
		select case mLev
			case 1, 2, 3, 4
			    fSecCode=Ucase(Request.Form("SecCode"))
			    If (SecImage<2) OR (SecImage>1 AND DoSecImage(fSecCode) = 1) Then
				  Call DoCookies(Request.Form("SavePassword"))
				  strLoginStatus = 1
				  closeAndGo(sScript)
				Else
				  strLoginStatus = 0
				End If
			case else
				strLoginStatus = 0
		end Select
	case "logout"
		Call ClearCookies()
		Session.Contents.RemoveAll()
		'Session.Abandon()
		strSql = "DELETE FROM " & strTablePrefix & "ONLINE WHERE UserIP='" & request.ServerVariables("REMOTE_ADDR") & "'"
		executeThis(strSql)
		closeAndGo("default.asp")
  end select
end if

':: check login status - member of guest
chkLoginStatus()

if mLev = 0 and strAuthType = "nt" and curpagetype <> "register" then
	  closeAndGo("register.asp?mode=register")
end if

':: populate member groups array
'dim arrAppPerms(), arrGroups()
bldArrUserGroup()

'this populates the arrCurOnline() array with online users membername and IP
buildOnlineUsersArray() 

'response.Write("<b>" & strDBNTusername & "</b><br>")
'response.Write("<br>LastHereDate: " & Session.Contents(strCookieURL & "last_here_date") & ":<br>")
if IsEmpty(Session(strCookieURL & "last_here_date")) or IsNull(Session(strCookieURL & "last_here_date")) or trim(Session(strCookieURL & "last_here_date")) = "" then
'response.Write("<br>LastHereDate: empty :<br>")	
Session.Contents(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTUserName)
else
'response.Write("<br>LastHereDate: not empty :<br>")
'refresh the session variable
Session.Contents(strCookieURL & "last_here_date") = Session.Contents(strCookieURL & "last_here_date")
end if
'response.Write("<br>LastHereDate: " & Session.Contents(strCookieURL & "last_here_date") & ":<br>")

':: build array for app access
bldArrAppAccess()
	  
if trim(curpagetype) <> "" and curpagetype <> "home" and curpagetype <> "register" and curpagetype <> "core" then
    if not chkApp(curpagetype,"USERS") then
	  closeAndGo("default.asp")
	end if
end if


Dim strOnlinePathInfo, strOnlineQueryString, strOnlineLocation
Dim strOnlineUser, strOnlineDate, strOnlineCheckInTime, strOnlineTimedOut
Dim strOnlineUsersCount, strOnlineGuestsCount, strOnlineMembersCount
Dim strOnlineGuestUserIP
' ******************************************************
' ADD HERE WHAT YOU WANT THE PREFIX OF YOUR COOKIE TO BE
' it will either be 'strCookieURL' or 'strUniqueID'
strTempCookieType = strCookieURL
' ******************************************************

strOnlinePathInfo = Request.ServerVariables("Path_Info")
strPgUrl = strOnlinePathInfo
strOnlineQueryString = "?" & Request.QueryString
if len(strOnlineQueryString) > 1 then
  'strPgUrl = strOnlinePathInfo & strOnlineQueryString
end if

' FIND OUT IF THEY ARE A GUEST, OR A USER
if strDBNTUserName = "" then
	strOnlineUser = txtGuest
else
	strOnlineUser = strDBNTUserName
end if
'
' Set Super admins IP to 0.0.0.0
if intIsSuperAdmin then
  arrWeb = split(left(strWebMaster,len(strWebmaster)-1),",")
  for olu = 0 to ubound(arrWeb)
    if lcase(strOnlineUser) = lcase(arrWeb(olu)) then
	  strOnlineUserIP = "0.0.0." & (olu + 1)
	end if
  next
  set arrWeb = nothing
else
  strOnlineUserIP = Request.ServerVariables("REMOTE_ADDR")
end if

' SET WHEN TO TIMEOUT THE USER
' DO THIS IN SECONDS
strOnlineDate = strCurDateString
strOnlineCheckInTime = strCurDateString

'  Count the current users online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [onlinecount] "
else
	strSqL = "SELECT count(UserID) onlinecount "
end if

strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
	on error resume next
	Err = 0
	Set rsOnline = my_Conn.Execute(strSql)
	If 0 <> Err Then
	  closeAndGo("site_setup.asp")
	end if
	on error goto 0
onlinecount = rsOnline("onlinecount")
strOnlineUsersCount = rsOnline("onlinecount")
set rsOnline = nothing

tmpOnlineUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
tmpAgent = split(tmpOnlineUserAgent," ")
strOnlineUserAgent = tmpAgent(0)

' Count the number of GUESTS online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Guests] "
else
	strSqL = "SELECT count(UserID) Guests "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) = '" & txtGuest & "' "

Set rsGuests = my_Conn.Execute(strSql)
Guests = rsGuests("Guests")
strOnlineGuestsCount = rsGuests("Guests")
Set rsGuests = nothing

' Count the number of MEMBERS online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Members] "
else
	strSqL = "SELECT count(UserID) Members  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) <> '" & txtGuest & "' "

Set rsMembers = my_Conn.Execute(strSql)
Members = rsMembers("Members")
strOnlineMembersCount = rsMembers("Members")
Set rsMembers = nothing


' Gets the current page title
CurPageTitle = UCase(Mid(CurPageType, 1, 1)) & Mid(CurPageType, 2, Len(CurPageType))

':::: these functions execute with each page load :::::::
  'check to see if it is a new day and run the once per day routine if needed
  OncePerDayChecks()
  'get current user skin
  getPageSkin(arrGroups(0,0))
  'custom_functions pageload call
  eachPageLoad()
':::: end page load functions :::::::::::::::::::::::::::
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- This page is generated by SkyPortal / SkyPortal.net <% =strForumTimeAdjust%>-->
<link rel="shortcut icon" href="<%= strHomeURL %>favicon.ico" type="image/x-icon">
<% getMetaTags() %>
<%

tmpPageTitle = ""
tmpPageTitle = strSiteTitle
if not CurPageType = "" and not uploadPg = true then
'tmpPageTitle = strSiteTitle & " | " & CurPageTitle
select case CurPageType
 case "forums"
	if not ChkString(Request("FORUM_Title"),"display") = " " then
	CurForumTitle =  " | " & ChkString(Request("FORUM_Title"),"display")
	else
	CurForumTitle = ""
	end if
	if not ChkString(Request("TOPIC_Title"),"display") = " " then
	CurTopicTitle = ChkString(Request("TOPIC_Title"),"display") & " - "
	else
	CurTopicTitle = ""
	end if
	tmpPageTitle = CurTopicTitle & strSiteTitle & " | " & CurPageTitle & CurForumTitle
 case "core"
    tmpPageTitle = PageTitle & " | " & strSiteTitle
 case else
	tmpPageTitle = CurPageTitle & " | " & strSiteTitle
 end select 
end if
%>
<title><% =tmpPageTitle%></title>
<% '######## RSS MODULE CODE ###### %>
<% If chkApp("rss","USERS") Then %>
   <link rel="alternate" type="application/rss+xml" title="<%=strSiteTitle%> News" href="<%=strHomeURL%>modules/rss/rss_news.asp" />
<% 
end if
%>

<script type="text/javascript">
// populate variables for use in the JS include files
 var js_welcome = "<%= txtWelcomeTo & " " & strSiteTitle %>";
 var txttype="<%= varBrowser %>";
 var js_collapse = "<%= txtCollapse %>";
 var js_expand = "<%= txtExpand %>";
 var js_none = "<%= txtNone %>";
 var js_member = "<%= txtMember %>";
 var js_admin = "<%= txtAdmin %>";
 
//  month/day arrays
 var js_months_lng=new Array("<%= txtJanuary %>", "<%= txtFebruary %>", "<%= txtMarch %>", "<%= txtApril %>", "<%= txtMay %>", "<%= txtJune %>", "<%= txtJuly %>", "<%= txtAugust %>", "<%= txtSeptember %>", "<%= txtOctober %>", "<%= txtNovember %>", "<%= txtDecember %>");
 var js_days_lng=new Array("<%= txtSunday %>", "<%= txtMonday %>", "<%= txtTuesday %>", "<%= txtWednesday %>", "<%= txtThursday %>", "<%= txtFriday %>", "<%= txtSaturday %>", "<%= txtSunday %>");
 
// pop-up calendar items
 var js_calendar = "<%= txtCalendar %>"
 var js_frm = "<%= txtForm %>"
 var js_frmfld = "<%= txtFrmFld %>"
 var js_notfnd = "<%= txtNotFound %>"
 var yxLinks=new Array("[<%= txtClose %>]", "[<%= txtClear %>]");
 
 var jsUniqueID = "<%= strUniqueID %>"

</script>
<script type="text/javascript" src="includes/scripts/cal2.js"></script>
<script type='text/javascript'> 
 // preload min-max images
  var mmImages = new Array(4);
  mmImages[0] = "Themes/<%= strTheme %>/icon_max.gif";
  mmImages[1] = "Themes/<%= strTheme %>/icon_min.gif";
  mmImages[2] = "Themes/<%= strTheme %>/icon_max1.gif";
  mmImages[3] = "Themes/<%= strTheme %>/icon_min1.gif";
  
  function jumpToPage(s) {if (s.selectedIndex != <%=mypage-1%>) top.location.href = s.options[s.selectedIndex].value;return 1;}
  
<%spThemeHeader_javascript()%>
</script>
<script type="text/javascript" src="modules/custom_scripts.js"></script>
<script type="text/javascript" src="includes/scripts/core.js"></script>
<script type='text/javascript' src="includes/scripts/menu_com.js"></script>
	<script type="text/javascript" src="includes/scripts/prototype.js"></script> 
	<script type="text/javascript" src="includes/scripts/effects.js"></script>
	<script type="text/javascript" src="includes/scripts/window.js"></script>

<% if hasEditor = true and editorType = "tinymce" and strAllowHtml = 1 then %>
<script language="javascript" type="text/javascript" src="tiny_mce/tiny_mce.js"></script>
<% End If %>
<!--#INCLUDE file="fp_custom.asp" -->
<!--#include file="includes/inc_editor.asp"-->
<% spThemeHeader_style() %>
<!--#include file="includes/inc_header.asp"-->
<% navBarTop() %>
<!--#include file="includes/inc_nav_bar.asp"-->
<% navBarBottom() %>
<%
'response.Flush()

if not uploadPg = true then	
  select case Request.Form("Method_Type")
	case "login"
	  response.Write("<br /><center><div style=""width: 500px"">")
	  spThemeblock1_open("1")
	  response.Write("<table border=""0""><tr><td width=""500"" align=""center"">")
	  response.Write("<p align=""center""><br /><br />")
	  if strLoginStatus = 0 then %>
	    <%= txtBadLogin1 %>
	    <% If SecImage >1 Then%>
	    <%= txtBadSecCode %>
	    <%End If%> 
	    <%= "&nbsp;" & txtWereInc %>.</p>
	    <p align="center">
	    <%= txtPlsTryAgain %>&nbsp;<%= txtOr %>&nbsp;<a href="policy.asp"><u><%= lcase(txtRegister) %></u></a>&nbsp;<%= txtForAccnt %>.
<%    else %>
	    <%= txtLoginSuccess %>!</p>
	    <p align="center">
	    <%= txtThksForPart %>.
	  <meta http-equiv="Refresh" content="1; URL=<%= sScript %>">
<%    end if %>
	  </p><br /><br /><br /><br />
	  <p align="center">
	  <a href="<% =sScript %>"><%= txtContinue %></a></p>
	  <%
	  response.Write("</td></tr></table>")
	  spThemeblock1_close("1")
	  response.Write("</div></center>") %>
	  <br /><br /><br /><br />
	  <% 
      showDaPage = false
		closeAndGo("stop")
  end select
end if

If (mLev = 0 and strLockDown = 1) Then
  if (curPageType = "register" and strNewReg = 1) or (strEmail = 1 and curpagetype = "password") then
	' do nothing
  else
    showDaPage = false
    if strDBauth = "nt" then
	  closeAndGo("register.asp")
	else
	  lockDownLoginForm()
	  closeAndGo("stop")
	end if
  end if
End If

if strForumStatus = "down" and CurPageType = "forums" and NOT intIsSuperAdmin then 
	  showForumDown()
      showDaPage = false
	  closeAndGo("stop")
end if 
'if showDaPage then
%>
<!--#include file="includes/inc_ipgate.asp" -->
<% 'if showDaPage then %>