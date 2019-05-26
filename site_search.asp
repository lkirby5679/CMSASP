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
<!-- #INCLUDE FILE="config.asp" -->
<%
PageTitle = txtSitSrch
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = txtSitSrch
	PageLocation = "site_search.asp"
	CurPageInfo = "<a href=" & PageLocation & ">" & PageName & "</a>"

end function
%>
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #INCLUDE file="includes/inc_ADOVBS.asp" -->
<!-- #INCLUDE FILE="inc_top.asp" -->
<%
'search = ChkString(Request("search"), "SQLString")
search = chkString(Request("search"),"sqlstring")
qsearch=replace(search," ","+")
'show = chkString(Request("num"),"sqlstring")
show = 10
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr>
<td width="190" class="leftPgCol" valign="top">
<% 
  intSkin = getSkin(intSubSkin,1)
  menu_fp() 
%></td>
<td width="100%" class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtSitSrch & "|site_search.asp"
  arg2 = txtSrchRslts & " : " & search & ":javascript:;"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

spThemeBlock1_open(intSkin)
Response.Write("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
response.write "<tr><td align=""left""><br>"
 
'############################## Forum Search ####################
If chkApp("forums","USERS") Then

Dim iPageSize       
Dim iPageCount      
Dim iPageCurrent    
Dim strOrderBy      
Dim strSQL          
Dim objPagingConn   
Dim objPagingRS     
Dim iRecordsShown   
Dim I

  strSQL = "select FORUM_ID from Portal_Topics where T_Subject like '%" & search & "%' or T_Message like '%" & search & "%' and T_Status=1" 

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
  reccount = 0
  if not objPagingRS.eof then
    do until objPagingRS.eof
      if chkForumAccess(strUserMemberID,objPagingRS("FORUM_ID")) then
	    reccount = reccount + 1
	  end if
	  objPagingRS.movenext
	loop
  end if

    objPagingRS.Close
    Set objPagingRS = Nothing

    strSQL = "select FORUM_ID from Portal_Reply where R_Message like'%" & search & "%' order by Reply_ID" 

    Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  if not objPagingRS.eof then
    do until objPagingRS.eof
      if chkForumAccess(strUserMemberID,objPagingRS("FORUM_ID")) then
	    reccount = reccount + 1
	  end if
	  objPagingRS.movenext
	loop
  end if
  objPagingRS.Close
  Set objPagingRS = Nothing
    'reccount = reccount + objPagingRS.recordcount
	%>
<center><span class="fTitle"><b><%= txtForums %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>
<% If reccount > 0 Then %>	
<center><a href="forum_search.asp?mode=DoIt&search=<%=qsearch%>&searchdate=0&Searchmember=0&SearchMessage=0&andor=phrase&forum=0">
<%= txtVSrchRslts %></a></center>
<br>
<% end if

response.Write("<hr />")
end if


'############################## Article Search ####################
If chkApp("article","USERS") Then
  strSQL = "select * from Article where Keyword like'%" & search & "%' or Summary like '%" & search & "%' or Content like '%" & search & "%' and show=1 order by Article_ID desc"

	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	reccount = objPagingRS.recordcount
	%>
	<center><span class="fTitle"><b><%= txtArticles %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>
	<br>
	<% 
 	If reccount > 0 Then
	%> 	
		<center><a href="article.asp?cmd=6&search=<%=search%>&submit1=Search&num=<%=show%>">
<%= txtVSrchRslts %></a></center><br>
	<%
	End If
	objPagingRS.Close
	Set objPagingRS = Nothing
	response.Write("<hr />")
end if 

'################# Picture Search Routine #############
If chkApp("pictures","USERS") Then
  strSQL = "select * from pic where Title like '%" & search & "%' or Keyword like'%" & search & "%' or description like '%" & search & "%' and show=1 order by pic_ID desc"

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

reccount = objPagingRS.recordcount
	%>
	<center><span class="fTitle"><b><%= txtPics %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br>
	<% 
If reccount > 0 Then
	%>
<center><a href="pic.asp?cmd=7&search=<%=search%>&submit1=Search&num=<%=show%>">
<%= txtVSrchRslts %></a></center>
<br>
	<%
End If

objPagingRS.Close
Set objPagingRS = Nothing
response.Write("<hr />")
end if 'strNavIcons

'############# Download Search ####################
If chkApp("downloads","USERS") Then
  strSQL = "select * from DL where Keyword like'%" & search & "%' or Description like '%" & search & "%' and show=1 order by DL_ID"

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
reccount = objPagingRS.recordcount
	%>
	<center><span class="fTitle"><b><%= txtDownloads %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br>
	<% 
If reccount > 0 Then
%> 	
<center><a href="dl.asp?cmd=7&search=<%=search%>&submit1=Search&num=<%=show%>">
<%= txtVSrchRslts %></a></center>
<br>
	<%
End If

objPagingRS.Close
Set objPagingRS = Nothing
	response.Write("<hr />")
end if 'strNavIcons

'#################### Link Search ###################
If chkApp("links","USERS") Then
  strSQL = "select * from links where Keyword like'%" & search & "%' or Description like '%" & search & "%' and show=1 order by link_ID"

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

reccount = objPagingRS.recordcount
	%>
	<center><span class="fTitle"><b><%= txtLinks %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br>
	<% 
If reccount > 0 Then
	%> 	
<center><a href="links.asp?cmd=7&search=<%=search%>&submit1=Search&num=<%=show%>">
<%= txtVSrchRslts %></a></center>
<br>
	<%
End If

objPagingRS.Close
Set objPagingRS = Nothing
	response.Write("<hr />")
end if 'strNavIcons

'################ CLASSIFIED SEARCH #######################
If chkApp("classifieds","USERS") Then
  strSQL = "select * from CLASSIFIED where Keyword like'%" & search & "%' or Description like '%" & search & "%' and show=1 order by CLASSIFIED_ID"

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

reccount = objPagingRS.recordcount
	%>
	<center><span class="fTitle"><b><%= txtClassifieds %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br>
	<% 
If reccount > 0 Then
	%>
<center><a href="classified.asp?cmd=7&search=<%=search%>&submit1=Search&num=<%=show%>">
<%= txtVSrchRslts %></a></center>
<br>
	<%
End If

objPagingRS.Close
Set objPagingRS = Nothing
	response.Write("<hr />")
end if 'strNavIcons

'############## Event Search Routine ################
If chkApp("events","USERS") Then
   strSQL = "select * from portal_events where Event_Title like'%" & search & "%' or Event_Details like '%" & search & "%' and private=0 order by Event_ID"

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

reccount = objPagingRS.recordcount
	%>
	<center><span class="fTitle"><b><%= txtEvents %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br>
	<% 
If reccount > 0 Then
	%>	
<center><a href="event_search.asp?search=<%=search%>&submit1=Search&num=<%=show%>">
<%= txtVSrchRslts %></a></center>
<br>
	<%
End If

objPagingRS.Close
Set objPagingRS = Nothing
	response.Write("<hr />")
end if 'strNavIcons

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
response.write "</td></tr>"
Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
<br>
</td>
</tr>
</table>
<!-- #INCLUDE FILE="inc_footer.asp" -->