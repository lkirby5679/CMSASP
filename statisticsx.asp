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

CurPageType = "core" %>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="lang/en/forum_core.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<%
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = txtSStats
	PageAction = txtViewing & "<br>" 
	PageLocation = "statisticsx.asp?" & strOnlineQueryString & ""
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<% if (mlev = 4 or mlev = 3) then


'dim intThisMonth

' Get the name of this file
'dim sScript
sScript = Request.ServerVariables("SCRIPT_NAME")

'set the date to today
'dim datToday
datToday = date()


' Check for valid month input
intThisMonth = trim(chkString(Request.QueryString("month"),"sqlstring"))


'set current month
if intThisMonth = "" then
	intThisMonth = month(datToday)
else
	intThisMonth = cint(intThisMonth)
end if

'constrain to only valid months
If intThisMonth < 1 OR intThisMonth > 12 Then
	intThisMonth = Month(datToday)
End If

intThisYear = trim(chkString(Request.QueryString("year"),"sqlstring"))

'set current year
If intThisYear = "" Then
	intThisYear = Year(datToday)
else
	intThisYear = cint(intThisYear)
End If
%>
<!--Sky Portal .net-->
<table cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td class="leftPgCol" valign="top">
	<%
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	%>
	<table width="100%" cellpadding="4" cellspacing="0">
  	<tr>
    <td align="left" class="tCellAlt1" width="100%">
      <span class="fTitle"><b>
      Jump To
      </b></span>
    </td>
  	</tr>
  	<tr>
    <td align="left" class="tCellAlt1">
    
	<FORM ACTION="<% =sScript%>" METHOD="GET" id="frmSelectMonth" name="frmSelectMonth">
		<SELECT NAME="month" onChange="this.form.submit()">
			<%for i = 1 to 12%>
			<OPTION VALUE="<%=i%>" <%if i = intThisMonth then Response.Write(" selected")%>><%= monthname(i,true)%>
			<%next%>
		</SELECT>
		<SELECT NAME="year" onChange="this.form.submit()">
			<%for i = -3 to 3%>
			<OPTION VALUE="<%= intThisYear + i%>" <%if (intThisYear + i) = intThisYear then Response.Write(" selected")%>><%= intThisYear + i%>
			<%next%>
		</SELECT>
	</FORM>
  <% If chkApp("forums","USERS") Then %>
    <li><a href="#postsday"><%= txtPstPerDay %></a></li>
    <li><a href="#postsmonth"><%= txtPstPerDMon %></a></li>
    <li><a href="#topicsday"><%= txtTopPerDay %></a></li>
    <li><a href="#topicsmonth"><%= txtTopPerMon %></a></li>
  <% End If %>
    <li><a href="#membersday"><%= txtMemPerDay %></a></li>
    <li><a href="#membersmonth"><%= txtMemPerMon %></a></li>
    
    </td>
  	</tr></table>
	<%
	spThemeBlock1_close(intSkin)
    menu_fp() %>
	</td>
	<td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = strSiteTitle & " " & txtStats & "|statistics.asp"
  arg2 = txtExStats & "|statisticsx.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>

<center><div style="width:500px;">
<%
%>
<br><br>
<%
spThemeTitle= strSiteTitle & " " & txtStats
spThemeBlock1_open(intSkin)
%>
<table class="tPlain" width="100%">

<% If chkApp("forums","USERS") Then %>
  <tr>
    <td align="left" class="tTitle" width="100%"><a name="postsday"></a>
      <span class="fTitle"><b><%= txtPstPerDay %></b></span>
    </td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt1">
    <%=DisplayDayPostCount%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0"></a></p>
    </td>
  </tr>

  <tr>
    <td align="left" class="tTitle" width="100%"><a name="postsmonth"></a>
      <span class="fTitle"><b><%= txtPstPerDMon %></b></span>
    </td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt1">
    <%=DisplayMonthPostCount%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0"></a></p>
    </td>
  </tr>

  <tr>
    <td align="left" class="tTitle" width="100%"><a name="topicsday"></a>
      <span class="fTitle"><b><%= txtTopPerDay %></b></span>
    </td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt1">
    <%=DisplayDayTopicsCount%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0"></a></p>
    </td>
  </tr>

  <tr>
    <td align="left" class="tTitle" width="100%"><a name="topicsmonth"></a>
      <span class="fTitle"><b><%= txtTopPerMon %></b></span>
    </td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt1">
    <%=DisplayMonthTopicsCount%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0"></a></p>
    </td>
  </tr>
<% End If %>
  <tr>
    <td align="left" class="tTitle" width="100%"><a name="membersday"></a>
      <span class="fTitle"><b><%= txtMemPerDay %></b></span>
    </td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt1">
    <%=DisplayDayMembersCount%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0"></a></p>
    </td>
  </tr>

  <tr>
    <td align="left" class="tTitle" width="100%"><a name="membersmonth"></a>
      <span class="fTitle"><b><%= txtMemPerMon %></b></span>
    </td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt1">
    <%=DisplayMonthMembersCount%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0"></a></p>
    </td>
  </tr></table>
<%
spThemeBlock1_close(intSkin)%></div></center>
    </td>
  </tr>
</table>
<!-- #include file="inc_footer.asp" -->
<%
function DisplayMonthPostCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	dim intYear, intMonth
	
	intYear = intThisYear
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
		<td valign="middle" align="center" width="15%"><span class="fSubTitle"><b>Months (<%=intYear%>)</b></span></td>
		<td valign="middle" align="center"><span class="fSubTitle"><b>Posts Count</b></span></td>
	</tr>
	<%
	intMonth = 1
	do while intMonth <= 12
		if intMonth = 1 then strMonth = "01"
		if intMonth = 2 then strMonth = "02"
		if intMonth = 3 then strMonth = "03"
		if intMonth = 4 then strMonth = "04"
		if intMonth = 5 then strMonth = "05"
		if intMonth = 6 then strMonth = "06"
		if intMonth = 7 then strMonth = "07"
		if intMonth = 8 then strMonth = "08"
		if intMonth = 9 then strMonth = "09"
		if intMonth = 10 then strMonth = "10"
		if intMonth = 11 then strMonth = "11"
		if intMonth = 12 then strMonth = "12"
		strSql = "SELECT COUNT(REPLY_ID) AS PostCount FROM " & strTablePrefix & "TOPICS INNER JOIN " & strTablePrefix & "REPLY ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID WHERE T_DATE LIKE '" & intYear & strMonth & "%' OR R_DATE LIKE '" & intYear & strMonth & "%'"
		'Response.write strsql
		'Response.end
		rs.Open strSql, my_Conn
	
	%>
	<tr>
		<td align="center" width="15%"><%= monthName(intMonth)%></td>
		<td align="center"><%= rs("PostCount")%></td>
	</tr>
	<%
		rs.close
		intMonth = intMonth + 1
	loop%>
</table>
<%
	set rs = nothing
end function

function DisplayDayPostCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	dim intYear, intMonth, strMonth, strDay
	
	intYear = intThisYear
	intMonth = intThisMonth
	intLastDay = getlastday(intMonth, intYear)
	if intMonth = 1 then strMonth = "01"
	if intMonth = 2 then strMonth = "02"
	if intMonth = 3 then strMonth = "03"
	if intMonth = 4 then strMonth = "04"
	if intMonth = 5 then strMonth = "05"
	if intMonth = 6 then strMonth = "06"
	if intMonth = 7 then strMonth = "07"
	if intMonth = 8 then strMonth = "08"
	if intMonth = 9 then strMonth = "09"
	if intMonth = 10 then strMonth = "10"
	if intMonth = 11 then strMonth = "11"
	if intMonth = 12 then strMonth = "12"
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
		<td valign="middle" align="center" width="15%"><span class="fSubTitle"><b><%=monthname(intMonth)%>&nbsp;<%=intYear%></b></span></td>
		<td valign="middle" align="center"><span class="fSubTitle"><b>Posts Count</b></span></td>
	</tr>
	<%
	intDay = 1
	do while intDay <= intLastDay
		if intDay = 1 then strDay = "01"
		if intDay = 2 then strDay = "02"
		if intDay = 3 then strDay = "03"
		if intDay = 4 then strDay = "04"
		if intDay = 5 then strDay = "05"
		if intDay = 6 then strDay = "06"
		if intDay = 7 then strDay = "07"
		if intDay = 8 then strDay = "08"
		if intDay = 9 then strDay = "09"
		if intDay >= 10 then strDay = cstr(intDay)

		strSql = "SELECT COUNT(REPLY_ID) AS PostCount FROM " & strTablePrefix & "TOPICS INNER JOIN " & strTablePrefix & "REPLY ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID WHERE T_DATE LIKE '" & intYear & strMonth & strDay & "%' OR R_DATE LIKE '" & intYear & strMonth & strDay & "%'"
		rs.Open strSql, my_Conn

	%>
	<tr>
		<td align="center" width="15%"><%= intDay %></td>
		<td align="center"><%= rs("PostCount")%></td>
	</tr>
	<%
		rs.close
		intDay = intDay + 1
	loop%>
</table>
<%
	set rs = nothing
end function

function DisplayMonthTopicsCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	dim intYear, intMonth
	
	intYear = intThisYear
	intMonth = intThisMonth
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
		<td valign="middle" align="center" width="15%"><span class="fSubTitle"><b>Months (<%=intYear%>)</b></span></td>
		<td valign="middle" align="center"><span class="fSubTitle"><b>Topics Count</b></span></td>
	</tr>
	<%

	intMonth = 1
	do while intMonth <= 12
		if intMonth = 1 then strMonth = "01"
		if intMonth = 2 then strMonth = "02"
		if intMonth = 3 then strMonth = "03"
		if intMonth = 4 then strMonth = "04"
		if intMonth = 5 then strMonth = "05"
		if intMonth = 6 then strMonth = "06"
		if intMonth = 7 then strMonth = "07"
		if intMonth = 8 then strMonth = "08"
		if intMonth = 9 then strMonth = "09"
		if intMonth = 10 then strMonth = "10"
		if intMonth = 11 then strMonth = "11"
		if intMonth = 12 then strMonth = "12"
		strSql = "SELECT COUNT(TOPIC_ID) AS PostCount FROM " & strTablePrefix & "TOPICS WHERE T_DATE LIKE '" & intYear & strMonth & "%'"
		'Response.write strsql
		'Response.end
		rs.Open strSql, my_Conn
	
	%>
	<tr>
		<td align="center" width="15%"><%= monthName(intMonth)%></td>
		<td align="center"><%= rs("PostCount")%></td>
	</tr>
	<%
		rs.close
		intMonth = intMonth + 1
	loop%>
</table>
<%
	set rs = nothing
end function

function DisplayDayTopicsCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	dim intYear, intMonth, strMonth, strDay
	
	intYear = intThisYear
	intMonth = intThisMonth
	intLastDay = getlastday(intMonth, intYear)
	if intMonth = 1 then strMonth = "01"
	if intMonth = 2 then strMonth = "02"
	if intMonth = 3 then strMonth = "03"
	if intMonth = 4 then strMonth = "04"
	if intMonth = 5 then strMonth = "05"
	if intMonth = 6 then strMonth = "06"
	if intMonth = 7 then strMonth = "07"
	if intMonth = 8 then strMonth = "08"
	if intMonth = 9 then strMonth = "09"
	if intMonth = 10 then strMonth = "10"
	if intMonth = 11 then strMonth = "11"
	if intMonth = 12 then strMonth = "12"
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
		<td valign="middle" align="center" width="15%"><span class="fSubTitle"><b><%=monthname(intMonth)%>&nbsp;<%=intYear%></b></span></td>
		<td valign="middle" align="center"><span class="fSubTitle"><b>Topics Count</b></span></td>
	</tr>
	<%
	intDay = 1
	do while intDay <= intLastDay
		if intDay = 1 then strDay = "01"
		if intDay = 2 then strDay = "02"
		if intDay = 3 then strDay = "03"
		if intDay = 4 then strDay = "04"
		if intDay = 5 then strDay = "05"
		if intDay = 6 then strDay = "06"
		if intDay = 7 then strDay = "07"
		if intDay = 8 then strDay = "08"
		if intDay = 9 then strDay = "09"
		if intDay >= 10 then strDay = cstr(intDay)

		strSql = "SELECT COUNT(TOPIC_ID) AS PostCount FROM " & strTablePrefix & "TOPICS WHERE T_DATE LIKE '" & intYear & strMonth & strDay & "%'"
		rs.Open strSql, my_Conn

	%>
	<tr>
		<td align="center" width="15%"><%= intDay %></td>
		<td align="center"><%= rs("PostCount")%></td>
	</tr>
	<%
		rs.close
		intDay = intDay + 1	
	loop%>
</table>
<%
	set rs = nothing
end function

function DisplayMonthMembersCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	dim intYear, intMonth
	
	intYear = intThisYear
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
		<td valign="middle" align="center" width="15%"><span class="fSubTitle"><b>Months (<%=intYear%>)</b></span></td>
		<td valign="middle" align="center"><span class="fSubTitle"><b>Members Count</b></span></td>
	</tr>
	<%
	intMonth = 1
	do while intMonth <= 12
		if intMonth = 1 then strMonth = "01"
		if intMonth = 2 then strMonth = "02"
		if intMonth = 3 then strMonth = "03"
		if intMonth = 4 then strMonth = "04"
		if intMonth = 5 then strMonth = "05"
		if intMonth = 6 then strMonth = "06"
		if intMonth = 7 then strMonth = "07"
		if intMonth = 8 then strMonth = "08"
		if intMonth = 9 then strMonth = "09"
		if intMonth = 10 then strMonth = "10"
		if intMonth = 11 then strMonth = "11"
		if intMonth = 12 then strMonth = "12"
		strSql = "SELECT COUNT(MEMBER_ID) AS MEMBERCOUNT FROM " & strTablePrefix & "MEMBERS WHERE M_DATE LIKE '" & intYear & strMonth & "%'"
		'Response.write strsql
		'Response.end
		rs.Open strSql, my_Conn
	
	%>
	<tr>
		<td align="center" width="15%"><%= monthName(intMonth)%></td>
		<td align="center"><%= rs("MEMBERCOUNT")%></td>
	</tr>
	<%
		rs.close
		intMonth = intMonth + 1
	loop%>
</table>
<%
	set rs = nothing
end function

function DisplayDayMembersCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	dim intYear, intMonth, strMonth, strDay
	
	intYear = intThisYear
	intMonth = intThisMonth
	intLastDay = getlastday(intMonth, intYear)
	if intMonth = 1 then strMonth = "01"
	if intMonth = 2 then strMonth = "02"
	if intMonth = 3 then strMonth = "03"
	if intMonth = 4 then strMonth = "04"
	if intMonth = 5 then strMonth = "05"
	if intMonth = 6 then strMonth = "06"
	if intMonth = 7 then strMonth = "07"
	if intMonth = 8 then strMonth = "08"
	if intMonth = 9 then strMonth = "09"
	if intMonth = 10 then strMonth = "10"
	if intMonth = 11 then strMonth = "11"
	if intMonth = 12 then strMonth = "12"
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
		<td valign="middle" align="center" width="15%"><span class="fSubTitle"><b><%=monthname(intMonth)%>&nbsp;<%=intYear%></b></span></td>
		<td valign="middle" align="center"><span class="fSubTitle"><b>Members Count</b></span></td>
	</tr>
	<%
	intDay = 1
	do while intDay <= intLastDay
		if intDay = 1 then strDay = "01"
		if intDay = 2 then strDay = "02"
		if intDay = 3 then strDay = "03"
		if intDay = 4 then strDay = "04"
		if intDay = 5 then strDay = "05"
		if intDay = 6 then strDay = "06"
		if intDay = 7 then strDay = "07"
		if intDay = 8 then strDay = "08"
		if intDay = 9 then strDay = "09"
		if intDay >= 10 then strDay = cstr(intDay)

		strSql = "SELECT COUNT(MEMBER_ID) AS MEMBERCOUNT FROM " & strTablePrefix & "MEMBERS WHERE M_DATE LIKE '" & intYear & strMonth & strDay & "%'"
		rs.Open strSql, my_Conn

	%>
	<tr>
		<td align="center" width="15%"><%= intDay %></td>
		<td align="center"><%= rs("MEMBERCOUNT")%></td>
	</tr>
	<%
		rs.close
		intDay = intDay + 1
	loop%>
</table>
<%
	set rs = nothing
end function

Function GetLastDay(intMonthNum, intYearNum)
	Dim dNextStart
	If CInt(intMonthNum) = 12 Then
		dNextStart = DateSerial(intYearNum,01,01)
	Else
		dNextStart = DateSerial(intYearNum,IntMonthNum+1,01)
	End If
	GetLastDay = Day(dNextStart - 1)
End Function



else%>
<p align="center">
<span class="fTitle">Sorry you are not allowed the view this page.</span></p>
<meta http-equiv="Refresh" content="2; URL=<% =chkString(Request.ServerVariables("HTTP_REFERER"), "refer") %>">
<p align="center">

<a href="<% =chkString(Request.ServerVariables("HTTP_REFERER"), "refer") %>">Back</a></p>
<!-- #include file=inc_footer.asp -->
<%end if%>
