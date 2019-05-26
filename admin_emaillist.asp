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
server.scripttimeout = 6000 
pgType = "SiteConfig"
%>
<!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/core_admin.asp" -->
<!--#INCLUDE file="includes/inc_adminvar.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<% 
page_name = ""
if request.form("page_name") <> "" then
  page_name = request.form("page_name")
end if
if page_name="compose" then
  'uploadPg = false
  hasEditor = true
  strEditorElements = "Message"
end if %>
<!--#INCLUDE file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#INCLUDE file="includes/inc_admin_functions.asp" -->
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
function selectAll(formObj, isInverse) 
{ 
with (formObj) 
{ 
for (var i=0;i < length;i++) 
{ 
fldObj = elements[i]; 
if(isInverse) 
{ 
if (fldObj.name != 'inverse') 
{ 
if (fldObj.name == 'selectall') 
fldObj.checked = false; 
else 
fldObj.checked = (fldObj.checked) ? false : true; 
} 
else fldObj.checked = true; 
} 
else 
{ 
fldObj.checked = true; 
if (fldObj.name == 'inverse') fldObj.checked = false; 
} 
} 
} 
}
// done hiding --> 
</script>
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
  arg1 = txtAdmin & "|admin_home.asp"
  arg2 = txtemUserEmailList & "|admin_emaillist.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
	<% spThemeBlock1_open(intSkin) %>
<%
if page_name = "" then
My_Sort = 0
My_Order = 0
My_Order2 = 1
if request.form("sortme") <> "" then
My_Sort = request.form("sortme")
end if
if request.form("order") <> "" then
My_Order = request.form("order")
end if
if request.form("order2") <> "" then
My_Order2 = request.form("order2")
end if

if My_Sort = 0 then
Sort_Name = txtemAllUsers
elseif My_Sort = 1 then
Sort_Name = txtemAdminOnly
elseif My_Sort = 2 then
Sort_Name = txtemModOnly
elseif My_Sort = 3 then
Sort_Name = txtemGenUsersOnly 
elseif My_Sort = 4 then
Sort_Name = txtemInactive6Mo
My_Last = DateToStr2(DateAdd("m", -6, now()))
elseif My_Sort = 5 then
Sort_Name = txtemInactive1Yr
My_Last = DateToStr2(DateAdd("yyyy", -1, now()))
elseif My_Sort = 6 then
Sort_Name = txtemNeverPosted
elseif My_Sort = 7 then
Sort_Name = txtemRefuseEmail
end if

if My_Order = 0 then
Order_Name = txtUsrName
elseif  My_Order =2 then
Order_Name = txtemRanking
elseif  My_Order =3 then
Order_Name = txtPosts 
elseif  My_Order =4 then
Order_Name = txtemStartDate
end if

if My_Order2 = 0 then
Order_By = "desc"
elseif My_Order2 = 1 then
Order_By = "asc"
end if

If Request.form("reccount") = "" Then
PageSize = 10	
Else
PageSize = Request.form("reccount")
End If
	
If PageSize = 1000000 then
PageSize2 = txtAll
Else
PageSize2 = PageSize
End If
	
If Request.form("pageno") = "" Then
intPage = 1	
Else
intPage = Request.form("pageno")
End If
%>
<% 
strSql = "SELECT * FROM " &strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> ''"

if My_Sort > 0 then
if My_Sort = 1 then
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 3
elseif My_Sort = 2 then
 strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 2
elseif My_Sort = 3 then 
 strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 1
elseif My_Sort = 6 then
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_POSTS = " & 0 
elseif My_Sort = 7 then      '---------------------------------------------------------------
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 1  
elseif My_Sort = 4 or My_Sort = 5 then 
 strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE < '" & My_Last & "'"
 end if
 end if
 
if My_Order = 2 then
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LEVEL " & Order_By
elseif My_Order = 3 then 
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_POSTS " & Order_By
elseif My_Order = 4 then 
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_DATE " & Order_By
else
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME " & Order_By
 end if
 
set rs = Server.CreateObject("ADODB.Recordset")
	
       rs.CursorLocation = 3	'adUseClient
      	rs.CursorType = 3		'adOpenStatic
      	rs.ActiveConnection = My_Conn
	rs.open strSql
	
	rs.PageSize = PageSize		
	rs.CacheSize = rs.PageSize
	intPageCount = rs.PageCount 
	intRecordCount = rs.RecordCount 

	If cint(intPage) > intPageCount Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
	
	If intRecordCount > 0 Then
		rs.AbsolutePage = intPage
		intStart = rs.AbsolutePosition
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rs.PageSize - 1)
		End if
	End If
	
'--------------------/paging

%>
<form method="post" action="admin_emaillist.asp" id="formEle">
<table border="0" cellpadding="0" cellspacing="10">
  <tr>
 <td nowrap>
<%=txtemSelUserGroup%><BR>
  
 <select name="sortme" size="1">
  <!-- <option value="<%=My_Sort%>" SELECTED>&nbsp;<% =Sort_Name%></option> -->
  <option value="0"<%= chkSelect(My_Sort,0) %>>&nbsp;<%=txtemAllUsers%></option>
  <option value="3"<%= chkSelect(My_Sort,cint(3)) %>>&nbsp;<%=txtemGenUsersOnly%></option>
  <option value="2"<%= chkSelect(cint(My_Sort),cint(2)) %>>&nbsp;<%=txtemModOnly%></option>
  <option value="1"<%= chkSelect(cint(My_Sort),cint(1)) %>>&nbsp;<%=txtemAdminOnly%></option>
  <option value="4"<%= chkSelect(cint(My_Sort),cint(4)) %>>&nbsp;<%=txtemInactive6Mo%></option>
  <option value="5"<%= chkSelect(cint(My_Sort),cint(5)) %>>&nbsp;<%=txtemInactive1Yr%></option>
  <option value="6"<%= chkSelect(cint(My_Sort),cint(6)) %>>&nbsp;<%=txtemNeverPosted%></option>
  <option value="7"<%= chkSelect(cint(My_Sort),cint(7)) %>>&nbsp;<%=txtemRefuseEmail%></option>
  </select>
  
 </td><td align="right">
<%=txtemOrderBy%><BR>
  
 <select name="order" size="1">
  <!-- <option value="<%=My_Order%>" SELECTED>&nbsp;<%=Order_Name%></option> -->
  <option value="0"<%= chkSelect(cint(My_Order),cint(0)) %>>&nbsp;<%=txtUsrName%></option>
  <option value="2"<%= chkSelect(cint(My_Order),cint(2)) %>>&nbsp;<%=txtemRanking%></option>
  <option value="3"<%= chkSelect(cint(My_Order),cint(3)) %>>&nbsp;<%=txtPosts%></option>
  <option value="4"<%= chkSelect(cint(My_Order),cint(4)) %>>&nbsp;<%=txtemStartDate%></option>
</select>
 
 </td>
<td align="center">
<%=txtemOrdering%><BR>
  
 <select name="order2" size="1">
  <!-- <option value="<%=My_Order2%>" SELECTED>&nbsp;<%=Order_By%></option> -->
  <option value=""<%= chkSelect(cint(My_Order2),cint(0)) %>>&nbsp;<%=txtemDescending%></option>
  <option value="1"<%= chkSelect(cint(My_Order2),cint(1)) %>>&nbsp;<%=txtemAscending%></option>
</select>&nbsp;<input type="submit" value="<%=txtGo%>" class="button">
 
 </td>
</tr>
 <tr>
 <td align="center">
<B><%=replace(txtemTotRecsInGroup,"[%marker_num%]",intRecordCount)%></b>
</td>
<td valign="top" align="right">
<%=txtemViewPage & " " & intPage%><BR>
<select name="pageno" size="1">
  <option value="1" SELECTED>&nbsp;1</option>
	<%	
	for counter = 2 to intPageCount
	   response.write"<option value=""" & counter & """" & chkSelect(cint(counter),cint(intPage)) & ">&nbsp;"& counter &"</option>"
	 
	Next
%>
	</select>
  </td><td valign="top" align="center">	
<%=txtemPerPage%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<BR>
<select name="reccount" size="1">
  <!-- <option value="<%= PageSize %>" SELECTED>&nbsp;<%= PageSize2 %></option> -->
  <option value="10" selected>&nbsp;10</option>
  <option value="25"<%= chkSelect(cint(PageSize),cint(25)) %>>&nbsp;25</option>
  <option value="50"<%= chkSelect(cint(PageSize),cint(50)) %>>&nbsp;50</option>
  <option value="100"<%= chkSelect(cint(PageSize),cint(100)) %>>&nbsp;100</option>
  <option value="200"<%= chkSelect(cint(PageSize),cint(200)) %>>&nbsp;200</option>
  <option value="500"<%= chkSelect(cint(PageSize),cint(500)) %>>&nbsp;500</option>
  <option value="1000"<%= chkSelect(cint(PageSize),cint(1000)) %>>&nbsp;1000</option>
  <option value="2000"<%= chkSelect(cint(PageSize),cint(2000)) %>>&nbsp;2000</option>
  <option value="1000000"<%= chkSelect(cint(PageSize),cLng(1000000)) %>>&nbsp;<%=txtAll%></option>
  </select>&nbsp;<input type="submit" value="<%=txtGo%>" class="button">
</td></tr>
</table></form>
  <form action="admin_emaillist.asp" method="post">
<hr width="90%" size="1" noshade>

<%         '------------------------------------
strSql = "SELECT * FROM " &strMemberTablePrefix & "SPAM WHERE ARCHIVE = '0'"
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql, my_Conn, 3
%>
<% if rsSP.EOF or rsSP.BOF then %>
<% else %>
<%=txtemSelectMessage%>
<select name="MSG" size="1">
<%
do until rsSP.EOF 
%>

  <option value="<% =rsSP("ID") %>">&nbsp;<% =Server.HTMLEncode(rsSP("Subject")) %></option>
  
<%
rsSP.MoveNext
%>
<%loop %>
</select>&nbsp;&nbsp;
<% end if 
set rsSP = nothing
%>

| <a href="admin_emailmanager.asp?mode=compose"><%=txtemCreateNewMessage%></a> | <a href="admin_emailmanager.asp"><%=txtemManageMessages%></a> |


<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
  <input type="hidden" name="page_name" value="compose">
  <td align=center>
  <input type="button" value="<%=txtemSelectVisible%>" onClick="selectAll(this.form,1)" class="button">&nbsp;
<input type="submit" name="action" value="<%=txtemSendMsgToSelected%>" class="button">
&nbsp;<input type="submit" name="action" value="<%=txtemSendMsgToAll%>" class="button">&nbsp;<input type="reset" name="Reset" value="- Reset -" class="button">
<BR><BR></td></tr>
 <td class="tCellAlt2">
 <table border="0" width="100%" cellspacing="1" cellpadding="4">
<tr align=center>
<td class="tTitle"><b><%=txtemMail%></b></td>
  <td class="tTitle"><b><%=txtUsrNam%></b></td>
  <td class="tTitle"><b><%=txtEmlAdd%></b></td>
  <td class="tTitle"><b><%=txtPosts%></b></td>
  <td class="tTitle"><b><%=txtemRanking%></b></td>  
  <td class="tTitle"><b><%=txtemRegistered%></b></td>
  <td class="tTitle">&nbsp;</td>  
</tr>

<% if rs.EOF or rs.BOF then %>
<tr>
  <td class="tCellAlt1" colspan="7"><b><%=txtNoMemFnd%></b></td>
</tr>
<% else %>
<%

'-------------------------paging
	For intRecord = 1 to rs.PageSize
	
	
MyRank = rs("M_LEVEL")
if MyRank = 1 then
MyRank = txtemGenUser
elseif MyRank = 2 then
MyRank = txtModerators
elseif MyRank = 3 then
MyRank = txtemAdministrator
end if
%>
<tr>
<td class="tCellAlt1" align="center">
<% '--------Does user want spam?
if rs("M_RECMAIL") = "1" then
%>
<B>X</B>
<% else %>
<input type="checkbox" name="ID" value="<% =rs("MEMBER_ID") %>"><input type="hidden" name="Mail_ALL" value="<% =rs("MEMBER_ID") %>">
<% end if %>
</td>
  <td class="tCellAlt1"><a href="cp_main.asp?cmd=8&member=<%=RS("MEMBER_ID")%>"><% =rs("M_NAME") %></a></td>
  <td class="tCellAlt1"><a href="mailto:<% =rs("M_EMAIL") %>"><% =rs("M_EMAIL") %></a></td>
  <td align=right class="tCellAlt1"><% =rs("M_POSTS") %></td><td align=right class="tCellAlt1"><% =MyRank %></td><td align=right class="tCellAlt1"><% =ChkDate(rs("M_DATE"))%>
</td><td class="tCellAlt1" align=center>
  <a href="cp_main.asp?cmd=10&mode=Modify&ID=<% =rs("MEMBER_ID") %>&name=<% =rs("M_NAME") %>"><img src="images/icons/icon_pencil.gif" alt="<%=txtemEditMember%>" border="0" hspace="0"></a>
  <a href="JavaScript:openWindow('pop_portal.asp?cmd=1&cid=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_trashcan.gif" alt="<%=txtemDelMember%>" border="0" hspace="0"></a>
  </td>
</tr>
<%
rs.MoveNext
If rs.EOF Then Exit for
Next
%>
<% end if %>
 </table>
 
 </td>
  </tr>
  <tr>
<td align="center"><BR><input type="button" value="<%=txtemSelectVisible%>" onClick="selectAll(this.form,1)" class="button">&nbsp;
<input type="submit" name="action" value="<%=txtemSendMsgToSelected%>" class="button">
&nbsp;<input type="submit" name="action" value="<%=txtemSendMsgToAll%>" class="button">&nbsp;<input type="reset" name="Reset" value="- Reset -" class="button"></td>
<BR></tr>
</table>

<%
elseif page_name="send" then
%>
<%
My_ID = request.form("id")
Mail_All = request.form("Mail_All")
if My_ID = "*" then
THISESCUELL = "select * from " & strMemberTablePrefix & "MEMBERS where MEMBER_ID in (" & Mail_All & ")"
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
else
THISESCUELL = "select * from " & strMemberTablePrefix & "MEMBERS where MEMBER_ID in (" & My_ID & ")"
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
end if
strSQL= THISESCUELL
set RS=Server.CreateObject("ADODB.Recordset")
RS.Open strSQL, my_Conn, 1, 3
rs.MoveFirst
while not rs.EOF
strRecipientsName = RS.Fields("M_NAME")
strRecipients = RS.Fields("M_email")
strFrom = strSender
strSubject = request.form("SUBJECT")
if request.form("html") = 1 then
 strMessage = request.form("MESSAGE")
 sendOutEmail strRecipients,strSubject,strMessage,0,1
else
 strMessage = request.form("MESSAGE")
 sendOutEmail strRecipients,strSubject,strMessage,2,0
end if
strArchive = request.form("ARCHIVE")
strL_DATE = strCurDateString
rs.MoveNext
wend 
%>
<h2><%=txtemMessageSent%></h2>
<p><center><%=txtemMsgSentToSelRecip%>
<br><br><a href="admin_emaillist.asp">Member Emaillist</a>
<br><br><a href="admin_home.asp">Admin Home</a></center></p>
<%
if request.form("save") <> "" then
strSubject = replace(request.form("SUBJECT"),"'","''")

if request.form("html") <> "" then
 strMessage = replace(request.form("MESSAGE"),"'","''")
else
 strMessage = replace(request.form("MESSAGE"),"'","''")
end if

	set conn = server.createobject ("adodb.connection")
	conn.open My_Conn
	conn.Execute "insert into " & strTablePrefix & "SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) values (" _
		& "'" & strSubject & "', " _
		& "'" & strMessage & "', " _ 
		& "'" & strL_DATE & "', " _		
		& "'" & request.form("ARCHIVE") & "')"

end if
%>

<%
'---------end sending
elseif page_name="compose" then
'-------------Start composing
%>
<% 
myMSG = Request.Form("MSG")
function getFormObject ()
if Request.ServerVariables("REQUEST_METHOD") = "GET" then
set getFormObject=Request.QueryString
else
set getFormObject=Request.Form
end if
end function
%>
<%
set oFormVars=GetFormObject()
if inStr(ucase(oFormVars("Action")),"SELECTED") > 0 then
selected=true
if oFormVars("ID") = "" then
  %>
 <h1><%=txtemNoRecipSel%></h1>
<%
response.end
end if
strSQL="select * from " & strMemberTablePrefix & "Members where MEMBER_ID in (" & request.form("id") & ")"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
else
selected=false
strSQL="select * from " & strMemberTablePrefix & "Members where MEMBER_ID in (" & request.form("Mail_All") & ")"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
end if

set RS=Server.CreateObject("ADODB.Recordset")
RS.Open strSQL, My_Conn, 1, 2

if myMSG <> "" then
strSql2 = "SELECT * FROM " & strTablePrefix & "SPAM WHERE ID =" & myMSG
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql2, My_Conn, 3
mySUBJECT = Server.HTMLEncode(rsSP("SUBJECT"))
myMESSAGE = rsSP("MESSAGE")
else                       
end if
Mail_All = request.form("Mail_All")
%>
<h4><%=txtSndMsg%></h4>
<form method="post" action="admin_emaillist.asp">
<input type="hidden" name="page_name" value="send">
<input type="hidden" name="Mail_All" value="<%= Mail_All %>">
<table class="grid" border="0" cellspacing="0" cellpadding="5">
<tr>
<td><%=txtSubject%>:</td><td><input class="textbox" type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
</tr>
<tr>
<td colspan="2"><%=txtMsg%>:</td>
</tr>
<tr>
<td colspan="2" align="center"><textarea class="textbox" name="Message" id="Message" cols="85" rows="25"><%= myMESSAGE %></textarea></td>
</tr>
</table>
<input type="checkbox" name="html" value="1">Use HTML?
<input type="checkbox" name="save" value="1"><%=txtemSaveThisMessage%>?&nbsp; 
 
 <select name="ARCHIVE" size="1">
  <option value="0" SELECTED>&nbsp;<%=txtemLiveList%></option>
  <option value="1">&nbsp;<%=txtemArchive%></option>
</select>
 

 &nbsp;<input type="Submit" value="Send" class="button">&nbsp;<input type="reset" class="button">
<%
if selected then
%>
<h4><i><%=txtemThisMsgSentToFllwngClnts%></i></h4>
<TABLE BORDER="1" class="grid" CELLSPACING="0" align="center" width="100%">
<TR>
<TH class="tTitle"><%=txtUsrName%></TH>
<TH class="tTitle"><%=txtEmlAdd%></TH>
</TR>
<%
On Error Resume Next
RS.MoveFirst
do while Not RS.eof
 %>
<TR VALIGN=TOP>
<td class="tCellAlt1"><input type="hidden" name="ID" value="<%=RS("MEMBER_ID")%>"><a href="cp_main.asp?cmd=8&member=<%=RS("MEMBER_ID")%>"><% =rs("M_NAME") %></a>&nbsp;</TD>
<td class="tCellAlt1"><% =rs("M_EMAIL") %>&nbsp;</TD>
</TR>
<%
RS.MoveNext
loop%>
</table>
<%
else
%>
<h4><i><%=txtemThisMsgSentToAllUsers%></i></h4>
<input type="hidden" name="ID" value="*"><%
end if
%></form>
<%
else 
  response.Write("page_name: " & page_name & "<br>")%>
<%=txtemUnknownAction%>
<%
response.end
end if
%><%
set rsSP = nothing
set rs = nothing
 spThemeBlock1_close(intSkin) %>
 </td></tr></table>
<!--#INCLUDE file="inc_footer.asp" --><% else %><%Response.Redirect "admin_login.asp?target=admin_emaillist.asp" %><% end iF %>