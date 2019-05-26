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

%>
<!--#include file="config.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<%
pageMode = trim(chkString(request.querystring ("pageMode"),"sqlstring"))
if Request.Querystring("mode") = "" then
	sMode = ""
else
	sMode = chkString(Request.QueryString("mode"),"sqlstring")
end if
if Request.Querystring("M_NAME") = "" then
	sName = ""
else
	sName = chkString(Request.QueryString("M_NAME"),"sqlstring")
end if
'response.Write(pageMode & "<br>") %>
<script type="text/javascript">
function ChangePage(pageNum){
document.location.href="pop_memberlist.asp?mode=<%=sMode%>&m_name=<%=sName%>&pageMode=<%=pageMode%>&pagesize=30&method=postsdesc&whichpage="+pageNum;
}
</script>
<%
if pageMode <> "" then
select case pageMode
	case "search"
%>
<script type="text/javascript">
function SearchMember(m_id, m_name) {
		pos = opener.document.SearchForm.SearchMember.length;
		if (pos<=1) {
		  opener.document.SearchForm.SearchMember.length +=1;
		  pos +=1;
		}  
		opener.document.SearchForm.SearchMember.options[pos-1].value = m_id;	
		opener.document.SearchForm.SearchMember.options[pos-1].text = m_name;
		opener.document.SearchForm.SearchMember.options[pos-1].selected = true;
}
</script>

<%	case "shoall" 
	  frm = chkString(Request.QueryString("frm"),"sqlstring")
	  sel = chkString(Request.QueryString("sel"),"sqlstring")
%>
<script type="text/javascript">
function AddMember(m_id, m_name) {

for (i=0; i<opener.document.<%= frm %>.<%= sel %>.length; i++) {
if (opener.document.<%= frm %>.<%= sel %>.options[i].value==m_id) {
	//user already added
	alert("<%= txtMemAlrAdd %>")
	return;
	}
}

		pos = opener.document.<%= frm %>.<%= sel %>.length;
		opener.document.<%= frm %>.<%= sel %>.length +=1;
		opener.document.<%= frm %>.<%= sel %>.options[pos].value = opener.document.<%= frm %>.<%= sel %>.options[pos-1].value;	
		opener.document.<%= frm %>.<%= sel %>.options[pos].text = opener.document.<%= frm %>.<%= sel %>.options[pos-1].text;
		opener.document.<%= frm %>.<%= sel %>.options[pos-1].value = m_id;	
		opener.document.<%= frm %>.<%= sel %>.options[pos-1].text = m_name;
		opener.document.<%= frm %>.<%= sel %>.options[pos-1].selected = true;
}
</script>
<%	case "allowmember"%>
<script type="text/javascript">
function AddAllowedMember(m_id, m_name) {

for (i=0; i<opener.document.PostTopic.AuthUsers.length; i++) {
if (opener.document.PostTopic.AuthUsers.options[i].value==m_id) {
	//user already added
	alert("<%= txtMemAlrAdd %>")
	return;
	}
}

		pos = opener.document.PostTopic.AuthUsers.length;
		opener.document.PostTopic.AuthUsers.length +=1;
		opener.document.PostTopic.AuthUsers.options[pos].value = opener.document.PostTopic.AuthUsers.options[pos-1].value;	
		opener.document.PostTopic.AuthUsers.options[pos].text = opener.document.PostTopic.AuthUsers.options[pos-1].text;
		opener.document.PostTopic.AuthUsers.options[pos-1].value = m_id;	
		opener.document.PostTopic.AuthUsers.options[pos-1].text = m_name;
		opener.document.PostTopic.AuthUsers.options[pos-1].selected = true;
}
</script>

<%	case "pmBan"%>
<script type="text/javascript">
function AddBannedMember(m_id, m_name) {

for (i=0; i<opener.document.PostTopic.BlockedUsers.length; i++) {
if (opener.document.PostTopic.BlockedUsers.options[i].value==m_id) {
	//user already added
	alert("<%= txtMemAlrAdd %>")
	return;
	}
}

		pos = opener.document.PostTopic.BlockedUsers.length;
		opener.document.PostTopic.BlockedUsers.length +=1;
		opener.document.PostTopic.BlockedUsers.options[pos].value = opener.document.PostTopic.BlockedUsers.options[pos-1].value;	
		opener.document.PostTopic.BlockedUsers.options[pos].text = opener.document.PostTopic.BlockedUsers.options[pos-1].text;
		opener.document.PostTopic.BlockedUsers.options[pos-1].value = m_id;	
		opener.document.PostTopic.BlockedUsers.options[pos-1].text = m_name;
		opener.document.PostTopic.BlockedUsers.options[pos-1].selected = true;
}
</script>

<%
end select

mypage = cLng(Request("whichpage"))
if mypage = 0 then
	mypage = 1
end if
mypagesize = cLng(Request.Querystring("pagesize"))
if mypagesize = 0 then
	mypagesize = 30
end if

If Request.QueryString("mode") = "search" then
mypagesize = 20
strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME LIKE '" & trim(chkstring(Request("M_NAME"), "sqlstring")) & "%' "

Set rsMembers = Server.CreateObject("ADODB.RecordSet")
rsMembers.open  strSql, my_conn, 3

if not (rsMembers.EOF or rsMembers.BOF) then  '## No categories found in DB
	rsMembers.movefirst
	rsMembers.pagesize = mypagesize

	maxpages = cint(rsMembers.pagecount)
end if
%>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
<tr><td>
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top" align="right"><b><%= txtPages %>:</b> &nbsp;&nbsp;</td>
        <td valign="top" align="right"><% Call Paging() %></td>
      </tr>
    </table>
<% end if %>
  </td></tr>
  <tr>
    <td>
<%
spThemeBlock1_open(intSkin)
%>
<table width="95%" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="center" class="tSubTitle"><%= txtMemName %>:</td>

<% If rsMembers.EOF or rsMembers.BOF then  '## No Members Found in DB %>
      <tr>
        <td align="center" class="tCellAlt1"><span class="fSubTitle"><b><%= txtNoMemFnd %></b></span>
        <p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p>
        </td>
      </tr>
<% Else 
	currMember = 0 %>
<%
	i = 0
	rsMembers.cacheSize = 30
	rsMembers.moveFirst
	rsMembers.pageSize = myPageSize
	maxPages = cint(rsMembers.pageCount)
	maxRecs = cint(rsMembers.pageSize)
	rsMembers.absolutePage = myPage
	howManyRecs = 0
	rec = 1
	do until rsMembers.Eof or rec = 31 
		if i = 1 then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if

memId = rsMembers("MEMBER_ID")
memName = ChkString(rsMembers("M_NAME"),"display")

select case pageMode
	case "search"
             Call selectMemSearch()
	case "shoall"
             Call selMemAllow()
	case "allowmember"
             Call selectMemAllow()
    case "pmBan"
             Call selectMemPmBan()
    case "pm"
             Call selectMemPm()
	case "all"
             Call selectAllMem()
	case "games"
             Call selectMemGames()
	case else
		response.write "ERROR!!!"
		response.end
end select

		currMember = rsMembers("MEMBER_ID")
		rsMembers.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop %>
<tr><td>
<% if maxpages > 1 then %><br>
    <table border=0 align="right">
      <tr>
        <td valign="top" align="right"><b><%= txtPages %>:</b> &nbsp;&nbsp;</td>
        <td valign="top" align="right"><% Call Paging() %></td>
      </tr>
    </table>
<% end if %>
  </td></tr>
      <tr>
        <td align="center" class="tCellAlt1"><br><p><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p></td>
      </tr>
<% end if%> 
</table>
<%
spThemeBlock1_close(intSkin)
%>
      </td></tr></table>
<% else
' - 
strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_STATUS, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE, " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE, " & strMemberTablePrefix & "MEMBERS.M_DATE "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
if mlev = 4 then
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
else
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
end if
if pageMode = "pmBan" then
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS=1"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL < 3"
end if
if pageMode = "pm" then
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS=1"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE=1"
end if
strSql = strSql & " ORDER BY M_POSTS DESC"

Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open  strSql, my_conn, 3

if not (rs.EOF or rs.BOF) then  '## No categories found in DB
	rs.movefirst
	rs.pagesize = mypagesize

	maxpages = cint(rs.pagecount)
end if
%>

<table border="0" width="95%" cellspacing="" cellpadding="0" align="center">
<% if maxpages > 1 then %>
  <tr>
    <td align="right"  width="100%">
    <table border=0 align="right">
      <tr>
        <td valign="top"><b><%= txtPages %>:</b> &nbsp;&nbsp;</td>
        <td valign="top"><% Call Paging() %></td>
      </tr>
    </table>
    </td>
  </tr>
<% end if %>
  <tr>
    <td>
<%
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
<tr><td>
<table cellspacing="0" cellpadding="0" width="100%">
      <tr>
        <td align="center" width="100%" class="tSubTitle"><span class="fSubTitle"><b><%= txtMemName %>:</b></span></td>
      </tr>
<% If rs.EOF or rs.BOF then  '## No Members Found in DB %>
      <tr>
        <td align="center" class="tCellAlt1"><span class="fSubTitle"><b><%= txtNoMemFnd %></b></span>
        <p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p>
        </td>
      </tr>
<% Else %>
<%	currMember = 0 %>
<%
	i = 0
	rs.cacheSize = 30
	rs.moveFirst
	rs.pageSize = myPageSize
	maxPages = cint(rs.pageCount)
	maxRecs = cint(rs.pageSize)
	rs.absolutePage = myPage
	howManyRecs = 0
	rec = 1
	do until rs.Eof or rec = 31 
		if i = 1 then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if

memId = rs("MEMBER_ID")
memName = ChkString(rs("M_NAME"),"display")
select case pageMode
	case "search"
             Call selectMemSearch()
	case "shoall"
             Call selMemAllow()
	case "allowmember"
             Call selectMemAllow()
    case "pmBan"
             Call selectMemPmBan()
	case "pm"
             Call selectMemPm()
	case "all"
             Call selectAllMem()
	case "games"
             Call selectMemGames()
	case else
		response.write txtError & "!!"
		response.end
end select

		currMember = rs("MEMBER_ID")
		rs.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop 
end if 
%>

</table></td></tr><tr><td>
 <form action="pop_memberlist.asp?mode=search&pageMode=<% = pageMode %>" method="post" name="SearchMembers">
<table cellpadding=0 cellspacing=0  width="100%">
 <tr class="tCellAlt1">
  <td><span class="fSubTitle"><b><%= txtSearch %>:</b>&nbsp;</span><input type="text" name="M_NAME" size="10">&nbsp;
  </td>
  <td>
   <INPUT type="submit" value="<%= txtSearch %>" id=submit1 name=submit1 border=0 width="40" height="25" class="button">
  </td>
 </tr> 
 </form> 
  <tr class="tCellAlt1">
    <td colspan="2" align="center" valign="top">        
	<a href="pop_memberlist.asp?mode=search&M_NAME=A&pageMode=<%= pageMode %>"><font  face=arial size=1>A</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=B&pageMode=<%= pageMode %>"><font  face=arial size=1>B</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=C&pageMode=<%= pageMode %>"><font  face=arial size=1>C</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=D&pageMode=<%= pageMode %>"><font  face=arial size=1>D</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=E&pageMode=<%= pageMode %>"><font  face=arial size=1>E</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=F&pageMode=<%= pageMode %>"><font  face=arial size=1>F</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=G&pageMode=<%= pageMode %>"><font  face=arial size=1>G</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=H&pageMode=<%= pageMode %>"><font  face=arial size=1>H</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=I&pageMode=<%= pageMode %>"><font  face=arial size=1>I</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=J&pageMode=<%= pageMode %>"><font  face=arial size=1>J</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=K&pageMode=<%= pageMode %>"><font  face=arial size=1>K</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=L&pageMode=<%= pageMode %>"><font  face=arial size=1>L</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=M&pageMode=<%= pageMode %>"><font  face=arial size=1>M</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=N&pageMode=<%= pageMode %>"><font  face=arial size=1>N</font></a><br>
	<a href="pop_memberlist.asp?mode=search&M_NAME=O&pageMode=<%= pageMode %>"><font  face=arial size=1>O</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=P&pageMode=<%= pageMode %>"><font  face=arial size=1>P</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=Q&pageMode=<%= pageMode %>"><font  face=arial size=1>Q</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=R&pageMode=<%= pageMode %>"><font  face=arial size=1>R</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=S&pageMode=<%= pageMode %>"><font  face=arial size=1>S</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=T&pageMode=<%= pageMode %>"><font  face=arial size=1>T</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=U&pageMode=<%= pageMode %>"><font  face=arial size=1>U</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=V&pageMode=<%= pageMode %>"><font  face=arial size=1>V</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=W&pageMode=<%= pageMode %>"><font  face=arial size=1>W</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=X&pageMode=<%= pageMode %>"><font  face=arial size=1>X</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=Y&pageMode=<%= pageMode %>"><font  face=arial size=1>Y</font></a>
	<a href="pop_memberlist.asp?mode=search&M_NAME=Z&pageMode=<%= pageMode %>"><font  face=arial size=1>Z</font></a><br>
	</td>
  </tr>
</table></td></tr></table>
<%
spThemeBlock1_close(intSkin)%>
    </td>
  </tr>
<% if maxpages > 1 then %>
<tr><td>
        <table border=0 width="100%">
          <tr>
            <td valign="top" align="right"><b><%= txtPages %>:</b></td>
            <td valign="top" align="right"><% Call Paging() %></td>
          </tr>
        </table>
</td></tr>
<% end if %>
</table>
<% end if
sub selectMemSearch()%>
      <tr>
        <td class="<% =CColor %>"><img src="<%= strHomeUrl %>Themes/<%= strTheme %>/icons/arrow1.gif" border="0">&nbsp;
        	<a href="javascript:void(0)" onClick="SearchMember('<%=memID%>', '<% =memName %>'); window.close()" title="<%= txtAddMem %>"><b><% =memName %></b></a>&nbsp;&nbsp;
			  <a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub selMemAllow()%>
      <tr>
        <td class="<% =CColor %>"><img src="<%= strHomeUrl %>Themes/<%= strTheme %>/icons/arrow1.gif" border="0">&nbsp;
        	<a href="javascript:void(0)" onClick="AddMember('<%=memID%>', '<% =memName %>');" title="<%= txtAddMem %>"><b><% =memName %></b></a>
        	  </td>
      </tr>
<%
end sub

sub selectMemAllow()%>
      <tr>
        <td class="<% =CColor %>"><img src="<%= strHomeUrl %>Themes/<%= strTheme %>/icons/arrow1.gif" border="0">&nbsp;
        	<a href="javascript:void(0)" onClick="AddAllowedMember('<%=memID%>', '<% =memName %>');" title="<%= txtAddMem %>"><b><% =memName %></b></a>&nbsp;&nbsp;
			  <a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub selectMemPmBan()%>
      <tr>
        <td class="<% =CColor %>"><a href="javascript:;" onClick="AddBannedMember('<%=memID%>', '<% =memName %>');" style="cursor:hand">&nbsp;
        <b><% =memName %></b></a>
        </td>
      </tr>
<%
end sub

sub selectMemPm()%>
      <tr>
        <td class="<% =CColor %>"><a onClick="opener.document.PostTopic.sendto.value+='<% =memName %>'; window.close()" style="cursor:hand">&nbsp;<img src="<%= strHomeUrl %>images/icons/pm.gif" width="11" height="17" title="<%= txtSndMsg %>" alt="<%= txtSndMsg %>" border="0">&nbsp;
        <b><% =memName %></b></a>&nbsp;&nbsp;<a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" title="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub selectAllMem()%>
      <tr>
        <td class="<% =CColor %>"><a onClick="opener.document.PostForm.member.value+='<% =memName %>'; window.close()" style="cursor:hand">&nbsp;<img src="<%= strHomeUrl %>images/icons/pm.gif" width="11" height="17" title="<%= txtSndMsg %>" alt="<%= txtSndMsg %>" border="0">&nbsp;
        <b><% =memName %></b></a>
        	  </td>
      </tr>
<%
end sub

sub selectMemGames()%>
      <tr>
        <td class="<% =CColor %>"><a onClick="opener.document.Bank.member.value+='<% =memName %>'; window.close()" style="cursor:hand">&nbsp;<img src="<%= strHomeUrl %>images/icons/pm.gif" width="11" height="17" alt="<%= txtSndMsg %>" title="<%= txtSndMsg %>" border="0">&nbsp;
        <b><% =memName %></b></a>&nbsp;&nbsp;<a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" title="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub Paging()
	if maxpages > 1 then
		if Request("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = chkString(Request("whichpage"),"sqlstring")
		end if
		Response.Write("<form name=""PageNum"" method=""post"" action=""pop_memberlist.asp?pageMode=" & pageMode & """>") & vbNewLine
		Response.Write("<select name=""whichpage"" size=""1"" onchange=""ChangePage(this.value)"">") & vbNewLine
		'Response.Write("<select name=""whichpage"" size=""1"" onchange=""submit()"">") & vbNewLine
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				Response.Write "<OPTION VALUE=""" & counter &  """>" & counter & vbNewLine
			else
				Response.Write "<OPTION SELECTED VALUE=""" & counter &  """>" & counter & vbNewLine
			end if
		next
		Response.Write("</select></form>")
	end if
end sub 
end if
%><!--#include file="inc_footer_short.asp" -->