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
'end if
isPrivateForum = 0
if CurPageInfoChk = "" then
	strOnlineLocation = txtUnkPg
else
  curPgLoc = CurPageInfo()
  if CurPageType = "forums" and isnumeric(strRqForumID) and trim(strRqForumID) <> "" then
  'response.Write("strRqForumID " & strRqForumID)
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS FROM " & strTablePrefix & "FORUM WHERE FORUM_ID = " & strRqForumID & " AND F_PRIVATEFORUMS <> 0"

	set rsPrf = my_Conn.execute(strSql)
	if not (rsPrf.BOF and rsPrf.EOF) then
		isPrivateForum = 1
	else
		isPrivateForum = 0
	end if
	rsPrf.Close
	set rsPrf = nothing
  end if
	if not curPgLoc = "" then
		if isPrivateForum = 1 then
			strOnlineLocation = txtPvtPg
		else
			strOnlineLocation = replace(curPgLoc,"&#59;",";")
		end if
	else 
	  strOnlineLocation = txtHidPg
	end if 
end if

strOnlineUser = OnlineSQLencode(strOnlineUser)
strOnlineLocation = OnlineSQLencode(strOnlineLocation)
strOnlineTimedOut = strOnlineCheckInTime - 1500      'time out the user after 25 minutes
if strOnlineUser <> txtGuest then
  strSql = "SELECT " & strTablePrefix & "ONLINE.UserID"
  strSql = strSql & " FROM " & strTablePrefix & "ONLINE "
  strSql = strSql & " WHERE " & strTablePrefix & "ONLINE.UserID='" & strOnlineUser & "'"
  isGuest = false
else
  strSql = "SELECT " & strTablePrefix & "ONLINE.UserID"
  strSql = strSql & " FROM " & strTablePrefix & "ONLINE "
  strSql = strSql & " WHERE " & strTablePrefix & "ONLINE.UserIP='" & strOnlineUserIP & "'"
  isGuest = true
end if
set rsWho =  my_Conn.Execute(strSql)

if rsWho.eof then
	  'THEY ARE A NEW USER SO INSERT THERE USERNAME
	  strSQL =  "INSERT INTO " & strTablePrefix & "ONLINE (UserID,UserIP,DateCreated,CheckedIn,LastChecked,M_BROWSE,UserAgent) VALUES ('"
	  strSql = strSQL & strOnlineUser & "','" & strOnlineUserIP & "','" & strOnlineDate & "','" & strOnlineCheckInTime & "','" & strOnlineCheckInTime & "','" & strOnlineLocation & "','" & strOnlineUserAgent & "')"
	  executeThis(strSql)
else
	' THEY ARE AN ACTIVE USER
	if not isGuest then
	' LETS UPDATE THE TABLE SO IT SHOWS THERE LAST ACTIVE VISIT
	strSql = "UPDATE " & strTablePrefix & "ONLINE SET M_BROWSE='" & strOnlineLocation & "', LastChecked='" & strOnlineCheckInTime & "', UserIP='" & strOnlineUserIP & "', UserAgent ='" & strOnlineUserAgent & "' WHERE UserID='" & strOnlineUser & "'"
	executeThis(strSql)
	else
	strSql = "UPDATE " & strTablePrefix & "ONLINE SET M_BROWSE='" & strOnlineLocation & "', LastChecked='" & strOnlineCheckInTime & "', UserAgent ='" & strOnlineUserAgent & "' WHERE UserIP='" & strOnlineUserIP & "'"
	executeThis(strSql)
	end if
end if
set rsWho = nothing
pop_pmToast()

'end if 'showDaPage
footerTop()
spThemeFooterBlock_open()%>
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
<td align="left" width="20"><a onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('<% =strHomeURL %>');"><img border="0" src="<% =strHomeURL %>Images/home.gif" title="<%= txtDefHomePG %>" alt="<%= txtDefHomePG %>" style="cursor: pointer;" /></a></td>
<td align="left" width="20"><a href="javascript:window.external.AddFavorite('<% =strHomeURL %>', '<% =strSiteTitle %>')"><img border="0" src="<% =strHomeURL %>Images/favorite.gif" title="<%= txtBkMkUs %>" alt="<%= txtBkMkUs %>" /></a></td>
<td align="center"><span class="fSmall"><a href="privacy.asp"><%= txtPrivacy %></a></span></td>
<td align="right">&nbsp;</td>
<td>&nbsp;</td>
<td align="center"><span class="fSmall"><% =strCopyright %></span></td>
<%'** START - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SkyPortal and SkyPortal.net  LICENSE AGREEMENT%>
<td align="right"><span class="fSmall">
<a href="http://www.SkyPortal.net" target="_blank" title="<%= txtPwrBy %>: SkyPortal Version <%= strWebSiteMVersion %>">SkyPortal.net</a></span></td>
<%'** END   - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SkyPortal and SkyPortal.net LICENSE AGREEMENT%>
<td width="10"><a href="#top"><img src="<% =strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" height="15" width="15" border="0" align="right" title="<%= txtTopPg %>" alt="<%= txtTopPg %>" /></a></td>
</tr></table>
<%spThemeFooterBlock_close()

If pageTimer = 1 Then
session.LCID = intPortalLCID
pageLoadTime = formatnumber((timer - startTime),3)
response.write "<br /><center><span class=""fSmall"">"
response.write txtPgLoad & " - " & pageLoadTime & "</span></center>"
end if

' LETS DELETE ALL INACTIVE USERS
SQL = "DELETE FROM " & strTablePrefix & "ONLINE WHERE LastChecked < '" & strOnlineTimedOut & "'"
on error resume next
my_Conn.execute(SQL)
on error goto 0

my_Conn.Close
set my_Conn = nothing

closeObjects()

spThemeEnd()
'end if 'showDaPage
%>
</body>
</html>