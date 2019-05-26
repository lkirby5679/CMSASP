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

img_new = "<img src=""themes/" & strTheme & "/new.gif"" border=""0"" alt="""" title=""New Items"" />"

function menu_fp()
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
if chkApp("events","USERS") then
eCnt = getCount("EVENT_ID","PORTAL_EVENTS","PENDING = 0 AND DATE_ADDED >= '" & Session(strCookieURL & "last_here_date") & "' AND PRIVATE = 0")
If eCnt = 0 Then eUrl = "events.asp" else eUrl = "events.asp?mode=newEvents" end if
'Pending Events count
PTcnt = PTcnt + getCount("EVENT_ID",strTablePrefix & "EVENTS","PENDING=1") 
end if

'::::::::::::::::::::::: Start the menu HTML ::::::::::::::::::::::::::::::
spThemeTitle= txtMenu
'spThemeTitle = spThemeTitle & " [" & intSkin & "]"
spThemeBlock1_open(intSkin)

defaultMenu()
%>

<table>
<tr><td width="100%"><hr /></td></tr>
<% if mLev >= 1 then
strSql = "SELECT " & strTablePrefix & "TOTALS.U_COUNT "
strSql = strSql & " FROM " & strTablePrefix & "TOTALS"
set rs1 = my_Conn.Execute(strSql)
Users = rs1("U_COUNT")
rs1.Close
set rs1 = nothing
%>
<tr><td width="100%"><span class="fSmall"><a href="members.asp"><%= txtMembers %>: <% =Users%></a></span></td></tr>
<% End If %>
<tr><td width="100%"><a href="active_users.asp"><span class="fSmall"><%= txtActvUsrs %>: <br /><%=strOnlineMembersCount & " " & txtMembers & " " & txtAnd & " " & strOnlineGuestsCount & " " & txtGuests %></span></a></td></tr></table>
<% 
spThemeBlock1_close(intSkin)
end function
%>