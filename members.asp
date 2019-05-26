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

CurPageType="core"
cnter = 0
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<%
PageTitle = txtMbrLst
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = txtMbrLst
	PageAction = txtViewing & "<br>" 
	PageLocation = "members.asp?" & strOnlineQueryString & ""
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<script type="text/javascript">
function ChangePage(){
	document.PageNum.submit();
}
</script>
<%
if mLev = 0 then
  closeAndGo("default.asp")
end if

If chkApp("forums","USERS") Then
  hasForums = true
else
  hasForums = false
end if

Dim	srchUName
Dim	srchFName
Dim	srchLName
Dim	srchState
Dim srchInitial
strPageSize = 15

Function sGetColspan(lIN, lOUT)	
if hasForums Then lOut = lOut + 2	
If (mlev = 4 or mlev = 3) then lOut = lOut + 1	
If mlev = 4 then lOut = lOut + 1		
If lOut > lIn then
	sGetColspan = lIN
Else
	sGetColspan = lOUT	
End If
End Function

if trim(chkString(Request.QueryString("method"),"SQLString")) <> "" then
	SortMethod = trim(chkString(Request.QueryString("method"),"SQLString"))
end if
SearchName = trim(chkString(Request("M_NAME"),"SQLString"))
if SearchName = "" then
	SearchName = trim(chkString(Request.Form("M_NAME"),"SQLString"))
end if
 
srchUName = trim(chkString(Request("UserName"),"SQLString"))
srchFName = trim(chkString(Request("FirstName"),"SQLString"))
srchLName = trim(chkString(Request("LastName"),"SQLString"))
srchInitial = trim(chkString(Request("INITIAL"),"SQLString"))
 
if IsNumeric(srchUName) <> True then srchUName = "1"
if IsNumeric(srchFName) <> True then srchFName = "0"
if IsNumeric(srchLName) <> True then srchLName = "0"
if IsNumeric(srchInitial) <> True then srchInitial = "0"
 
mypage = trim(chkString(request("whichpage"),"SQLString"))

if ((Trim(mypage) = "") Or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = CInt(mypage)

' Paging Variables
dim scriptname, intPagingLinks, strQS
scriptname = request.servervariables("script_name")
intPagingLinks = 6 ' ## Number of links per page...
strQS = "&initial=" & srchInitial &_
	"&M_NAME=" & SearchName &_
	"&mode=search" &_
	"&method=" & SortMethod

'New Search Code
If Request("mode") = "search"  and (srchUNAME = "1" or srchFNAme = "1" or srchLName = "1" or srchState = "1" or srchInitial = "1" ) then 

	strSqlm = "SELECT " & strMemberTablePrefix & "MEMBERS.*"
	strSqlm = strSqlm & " FROM " & strMemberTablePrefix & "MEMBERS "
if Request.querystring("link") <> "sort" then
	whereSql = " WHERE (" 
	tmpSql = ""
	if srchState <> "1" then
		if srchUName = "1" and srchInitial = "0" then
		tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_NAME LIKE '%" & SearchName & "%' OR "
		tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_USERNAME LIKE '%" & SearchName & "%'"
		end if
		if srchFName = "1" and srchInitial = "0" then
			if srchUName = "1" then
					tmpSql = tmpSql & " OR "
			end if
			tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME LIKE '%" & SearchName & "%'"
		end if
		if srchLName = "1" and srchInitial = "0" then
			if srchFName = "1" or srchUName = "1" then 
				tmpSql = tmpSql & " OR "
			end if
			tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_LASTNAME LIKE '%" & SearchName & "%' "
		end if
		if srchInitial = "1" then 
		  tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_NAME LIKE '" & SearchName & "%' OR "
		  tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_USERNAME LIKE '" & SearchName & "%'"
		end if
	else
		tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_STATE LIKE '%" & SearchName & "%'"
	end if
	whereSql = whereSql & tmpSql &")"
	Session(strCookieURL & "where_Sql") = whereSql
end if	

	if Session(strCookieURL & "where_Sql") <> "" then
		whereSql = Session(strCookieURL & "where_Sql")
	else
		whereSql = ""
	end if
	strSqlm = strSqlm & whereSql
else

	strSqlm = "SELECT " & strMemberTablePrefix & "MEMBERS.*"
	strSqlm = strSqlm & " FROM " & strMemberTablePrefix & "MEMBERS "
	if mlev < 4 then
		strSqlm = strSqlm & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
		strSqlm = strSqlm & " AND " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
	else
		'strSqlm = strSqlm & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	end if
end if
select case SortMethod
	case "nameasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "namedesc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME DESC"
	case "levelasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_TITLE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "leveldesc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_TITLE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastpostdateasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastpostdatedesc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastheredateasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastheredatedesc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "dateasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_DATE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "datedesc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_DATE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "countryasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_COUNTRY ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "countrydesc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_COUNTRY DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "postsasc"
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_POSTS ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case else
		strSqlm = strSqlm & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LEVEL DESC, " & strMemberTablePrefix & "MEMBERS.M_POSTS DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
end select

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		OffSet = CInt((mypage - 1) * strPageSize)
		strSqlm = strSqlm & " LIMIT " & OffSet & ", " & strPageSize & " "
	end if

	' - Get the total pagecount 
	strSqlm2 = "SELECT COUNT(" & strMemberTablePrefix & "MEMBERS.MEMBER_ID) AS PAGECOUNT "
	strSqlm2 = strSqlm2 & " FROM " & strMemberTablePrefix & "MEMBERS " 
	if mlev = 4 then
		strSqlm2 = strSqlm2 & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
	else
		strSqlm2 = strSqlm2 & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	end if


	set rsCountx = my_Conn.Execute(strSqlm2)
	if not rsCountx.eof then
		maxpages = (rsCountx("PAGECOUNT")  \ strPageSize )
			if rsCountx("PAGECOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		maxRecs = cint(strPageSize) * maxPages
	else
		maxpages = 0
	end if 

	rsCountx.close
	
	set rsd = Server.CreateObject("ADODB.Recordset")

	rsd.open  strSqlm, my_Conn, 3
	
	if not (rsd.EOF or rsd.BOF) then
		rsd.movefirst
	end if
 
else 'end MySql specific code

	Set rsd = Server.CreateObject("ADODB.RecordSet")
	rsd.cachesize=20

	rsd.open strSqlm, my_conn, 3

	if not (rsd.EOF or rsd.BOF) then  '## No members found in DB
		rsd.movefirst
		rsd.pagesize = strPageSize
		rsd.cacheSize = strPageSize
		maxPages = cint(rsd.pageCount)
		maxRecs = cint(rsd.pageSize)
		rsd.absolutePage = myPage
		maxpages = cint(rsd.pagecount)
	end if
end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td class="leftPgCol" align="center" valign="top">
	<%
	intSkin = getSkin(intSubSkin,1)
	menu_fp() %>
	</td>
	<td class="mainPgCol" valign="top" width="100%">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtMbrLst & "|members.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<% if maxpages > 1 then %>
    <table border="0" align="center" width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top" width="50%" align="right"><b><%= txtGoToPg %>:</b> &nbsp;&nbsp;</td>
        <td valign="top"><% Call Paging2() %></td>
      </tr>
    </table>
<% end if %>

<%
spThemeTitle = txtSrchMem & ":"
spThemeBlock1_open(intSkin)%>
 <form action="members.asp?method=<%=SortMethod %>" method="post" name="SearchMembers">
 <center><table class="tCellAlt0" cellpadding="0" cellspacing="0"><tr>
   <td valign="top"><b></b>  
  <div class="clsSpacer"></div>
  <input type="checkbox" name="UserName" value="1" <% if srchUName <> "0" then response.write "checked" %> ><%= txtUsrNam %>
  &nbsp;&nbsp;<input type="checkbox" name="FirstName" value="1" <% if trim(srchFName) <> "0" then response.write "checked" %> ><%= txtFstNam %>
  &nbsp;&nbsp;<input type="checkbox" name="LastName" value="1" <% if srchLName <> "0" then response.write "checked" %>><%= txtLstNam %>&nbsp;
  </td>
  <td><b><%= txtFor %>:</b>&nbsp;
  <br><input type="text" name="M_NAME" size="20" value="<%= SearchName %>"></td>
<input type="hidden" name="mode" value="search">
<input type="hidden" name="initial" value="<%= srchInitial %>">
  <td valign="middle" align="center"> &nbsp;
  <br>
    <input class="button" type="submit" name="Submit" value=" Search ">
  </td>
 </tr> 
  <tr>
    <td colspan="3" align="center" valign="center" height="20"> 
	<% 
	arrAlpha = split(txtAlphabet,",")
	response.Write("<a href=""members.asp"">" & txtAll & "</a>&nbsp;")
	for xa = 0 to ubound(arrAlpha)
	response.Write("&nbsp;<a href=""members.asp?mode=search&M_NAME=" & arrAlpha(xa) & "&initial=1&method=" & SortMethod & """>" & arrAlpha(xa) & "</a>")
	next
	%><br>
	</td>
  </tr></table></center>
 </form>
<%spThemeBlock1_close(intSkin)%>
<br>

<%
spThemeBlock1_open(intSkin)
%>
    <table border="0" width="100%" cellspacing="1" cellpadding="3">
      <tr>
        <td align="center" class="tSubTitle">&nbsp;&nbsp;</td>
        <td align="left" class="tSubTitle"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "nameasc" then Response.Write("namedesc") else Response.Write("nameasc") end if %>"><b><%= txtMemName %></b></a></td>
        <td align="center" class="tSubTitle"><b><%= txtTitle %></b></td>
	<% If hasForums Then %>
        <td align="center" class="tSubTitle"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "postsdesc" then Response.Write("postsasc") else Response.Write("postsdesc") end if %>"><b><%= txtPosts %></b></a></td>
        <td align="center" class="tSubTitle"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "lastpostdatedesc" then Response.Write("lastpostdateasc") else Response.Write("lastpostdatedesc") end if %>"><b><%= txtLstPost %></b></a></td>
	<% End If %>
        <td align="center" class="tSubTitle"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "datedesc" then Response.Write("dateasc") else Response.Write("datedesc") end if %>"><b><%= txtMbrSnce %></b></a></td>
        <td align="center" class="tSubTitle"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "countryasc" then Response.Write("countrydesc") else Response.Write("countryasc") end if %>"><b><%= txtCntry %></b></a></td>
<% if mlev = 4 or mlev = 3 then %>
        <td align="center" class="tSubTitle"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "lastheredatedesc" then Response.Write("lastheredateasc") else Response.Write("lastheredatedesc") end if %>"><b><%= txtLstVst %></b></a></td>
<% end if %>
<% if mlev = 4 then %>
        <td align="center" class="tSubTitle"><b><%= txtOptions %></b></td>
<% end if %>
      </tr>
<% if rsd.EOF then  '## No Members Found in DB %>
      <tr>
        <td colspan="<%=sGetColspan(9, 5)%>" ><b><%= txtNoMemFnd %></b></td>
      </tr>
<% else %>
<%	currMember = 0 %>
<%
'	i = 0
	howManyRecs = 0
	rec = 1
	CColor = "tCellAlt2"
	do until rsd.Eof or rec = (strPageSize + 1)
		if CColor = "tCellAlt1" then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if
%>
      <tr class="<% =CColor %>">
        <td align=left>
<%	
  cnter = cnter + 1 %>
        <% if rsd("M_STATUS") = 0 then 
        response.Write "<img src=""images/icons/icon_profile_locked.gif"" title=""" & txtMemLckd & """ alt=""" & txtMemLckd & """ height=""15"" width=""15"" border=""0"" hspace=""0"">"
		else %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');"><img src="images/icons/icon_group.gif" title="<%= txtMbrCnct %>" alt="<%= txtMbrCnct %>" border="0" hspace="0" align="absmiddle"></a>
		<% 'memImg = ""
		end if %>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:100px; z-index:100; display:none; position:absolute; left:220px;">
<%  'cnter = 1
getMiniProfile() %>
<center><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');"><span class="fSmall"><%= txtClose %></span></a></center>
</div>

        </td>
        <td>
        <% strIMmsg = "View " & ChkString(rsd("M_NAME"),"display") & "'s profile" %>
		<a href="cp_main.asp?cmd=8&member=<% =rsd("MEMBER_ID") %>">
	  	<b><acronym title="<%= strIMmsg %>">
	  <%= displayName(ChkString(rsd("M_NAME"),"display"),trim(rsd("M_GLOW"))) %>
	  </acronym></b></a>
		</td>
        <td><%= getDonor_Level(rsd("MEMBER_ID")) %>
		<% =ChkString(getMember_Level(rsd("M_TITLE"), rsd("M_LEVEL"), rsd("M_POSTS")),"display") %></td>
	 <% If hasForums Then %>
        <td align=center>
<%		if IsNull(rsd("M_POSTS")) then %>
        -
<%		else %>
         <% =rsd("M_POSTS") %>
<%			if strShowRank = 2 or strShowRank = 3 then 
%>
        <br><% Response.write(getStar_Level(rsd("M_LEVEL"), rsd("M_POSTS"))) %>
<%			end if %>
<%		end if %>
        </td>
        <% if IsNull(rsd("M_LASTPOSTDATE")) or Trim(rsd("M_LASTPOSTDATE")) = "" then%>
        <td align=center nowrap>-</td>
        <% else %>
        <td align=center nowrap><% =ChkDate(rsd("M_LASTPOSTDATE")) %></td>
        <% end if %>
      <% End If %>
        <td align=center nowrap><% =ChkDate(rsd("M_DATE")) %></td>
        <td align=center><% =rsd("M_COUNTRY") %>&nbsp;</td>
<%		if mlev = 4 or mlev = 3 then %>
        <td align=center nowrap><% =ChkDate(rsd("M_LASTHEREDATE")) %></td>
<%		end if %>
<%		if mlev = 4 then %>
        <td align=center>
<%
  cnter = cnter + 1 %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');mwpHSs('formEle','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onMouseOver="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onMouseOut="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtMbrOpts %>" alt="<%= txtMbrOpts %>" border="0" hspace="0" align="absmiddle"></a>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:120px; z-index:100; display:none; position:absolute; right:50px;">
<table class="tPlain" width="110"><tr><td align="center" nowrap="nowrap">
<%  'cnter = 1
	response.Write("" & txtMbrOpts & ":<br>")
	if instr(strWebMaster,"" & lcase(rsd("M_NAME")) & ",") <> 0 and instr(strWebMaster,"" & lcase(strDBNTUserName) & ",") <> 0 then 
		if rsd("M_STATUS") <> 0 then %>
			<a href="JavaScript:openWindow('pop_portal.asp?cmd=2&cid=<% =rsd("MEMBER_ID") %>')" title="<%= txtLock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_lock.gif" alt="<%= txtLock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%		else %>
          	<a href="JavaScript:openWindow('pop_portal.asp?cmd=3&cid=<% =rsd("MEMBER_ID") %>')" title="<%= txtUlock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_unlock.gif" alt="<%= txtUlock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%		end if
	else
	  if instr(strWebMaster,"" & lcase(rsd("M_NAME")) & ",") = 0 then %>
<%		if rsd("M_STATUS") <> 0 then %>
          <a href="JavaScript:openWindow('pop_portal.asp?cmd=2&cid=<% =rsd("MEMBER_ID") %>')" title="<%= txtLock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_lock.gif" alt="<%= txtLock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%		else %>
          <a href="JavaScript:openWindow('pop_portal.asp?cmd=3&cid=<% =rsd("MEMBER_ID") %>')" title="<%= txtUlock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_unlock.gif" alt="<%= txtUlock %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%		end if %>
<%	  end if 
	end if
	if instr(strWebMaster,"" & lcase(rsd("M_NAME")) & ",") <> 0 then %>
<%		if intIsSuperAdmin = 1 then %>
			<a href="cp_main.asp?cmd=10&mode=Modify&ID=<% =rsd("MEMBER_ID") %>&name=<% =ChkString(rsd("M_NAME"),"urlpath") %>" title="<%= txtEdit %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_pencil.gif" alt="<%= txtEdit %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%		end if %>
<%	else %>
		<a href="cp_main.asp?cmd=10&mode=Modify&ID=<% =rsd("MEMBER_ID") %>&name=<% =ChkString(rsd("M_NAME"),"urlpath") %>" title="<%= txtEdit %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_pencil.gif" alt="<%= txtEdit %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%	end if
	if instr(strWebMaster,"" & lcase(rsd("M_NAME")) & ",") <> 0 then %>
<%		'do nothing %>
<%	else %>
        <a href="JavaScript:openWindow('pop_portal.asp?cmd=1&cid=<% =rsd("MEMBER_ID") %>')" title="<%= txtDel %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>"><img src="images/icons/icon_trashcan.gif" alt="<%= txtDel %>&nbsp;<% =ChkString(rsd("M_NAME"),"display") %>" border="0" hspace="0"></a>
<%	end if 
	if rsd("M_LEVEL") = 1 then %>
        <a href="cp_main.asp?mode=Moderator&action=add&ID=<% =rsd("MEMBER_ID") %>" title="<%= txtMakMod %>"><img src="images/icons/icon_mod.gif" alt="<%= txtMakMod %>" border="0" hspace="0"></a>
<%  Elseif  rsd("M_LEVEL") = 2 then%>
		<a href="cp_main.asp?mode=Moderator&action=del&ID=<% =rsd("MEMBER_ID") %>" title="<%= txtRemMod %>"><img src="images/icons/icon_delmod.gif" alt="<%= txtRemMod %>" border="0" hspace="0"></a>			
<%  End If
	if varBrowser = "ie" then
		if rsd("M_GLOW") <> "" then %>
          <a href="javascript:;" title="<%= txtRemGlo %>"><img src="images/icons/icon_color.gif" onClick="openWindow('pop_glow.asp?cmd=2&id=<% =rsd("MEMBER_ID") %>')" alt="<%= txtRemGlo %>" border="0" hspace="0"></a>
	<%  Else %>
		  <a href="javascript:;" title="<%= txtAddEdtGlo %>"><img src="images/icons/icon_color.gif" onClick="openWindow('pop_glow.asp?cmd=1&id=<% =rsd("MEMBER_ID") %>')" alt="<%= txtAddEdtGlo %>" border="0" hspace="0"></a>			
	<%  End If
	end if %>
		</td></tr><tr><td align="center"><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1'); mwpHSs('formEle','1');"><span class="fSmall"><%= txtClose %></span></a></td></tr></table>
</div>
</td>
<%end if %>
      </tr>
<%		currMember = rsd("MEMBER_ID")
		rsd.MoveNext
		rec = rec + 1
	loop 
end if 
%>
  <tr>
    <td colspan="<%=sGetColspan(9, 5)%>">
<% if maxpages > 1 then %>
        <table border="0" width="100%">
          <tr>
            <td valign="top"><b><% =maxpages %>&nbsp;<%= txtPages %></b> &nbsp;&nbsp;</td>
            <td valign="top"><% Call Paging() %></td>
          </tr>
        </table>
<% else %>
        &nbsp;
<% end if %>
        </td>
      </tr>
    </table>
<%
spThemeBlock1_close(intSkin)%>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
sub Paging2()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = chkString(Request.QueryString("whichpage"),"sqlstring")
		end if
		if Request.QueryString("method") = "" then
			sMethod = "postsdesc"
		else
			sMethod = chkString(Request.QueryString("method"),"sqlstring")
		end if

		sScriptName = Request.ServerVariables("script_name")
		Response.Write("<form name=""PageNum"" action=""members.asp?method=" & SortMethod & "&Initial=" & initial & "mode=search&M_NAME=" & searchName & """>")
		Response.Write("<select name=""whichpage"" size=""1"" onchange=""ChangePage()"">")
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				Response.Write "<OPTION VALUE=""" & counter &  """>" & counter
			else
				Response.Write "<OPTION SELECTED VALUE=""" & counter &  """>" & counter
			end if
		next
		Response.Write("</select>")

	end if
end sub 

sub Paging()

	if (IsNumeric(intPagingLinks) = 0) AND (Trim(intPagingLinks) = "") then intPagingLinks = 10
	if (maxpages > 1) and (Trim(strQS) <> "") then
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""0"" valign=""top"" align=""center"">" & vbCrLf &_
			"<tr align=""center"">" & vbCrLf)
		if maxpages > 10 then
			Response.Write("<td>")
			Response.Write("<form method=""post"" name=""pagelist"" id=""pagelist"" action=""" & scriptname & "?n=0"& strQS & """>")
			Response.Write("<table cellpadding=""0"" cellspacing=""0"" border=""0"" align=""right""><tr><td><b>" & txtGoToPg & "</b>:&#160;</td><td>")
			Response.Write("<select name=""whichpage"" onchange=""jumpToPage(this)"" style=""font-size:10px;"">" & vbCrLf)
			Response.Write("<option value=""" & scriptname & "?whichpage=1" & strQS & """>&#160;-" & vbCrLf)
			pgeselect = ""
			if pgenumber = mypage then pgeselect = " selected"
			Response.Write("<option value=""" & scriptname & "?whichpage=1" & strQS & """" & pgeselect & ">1" & vbCrLf)
			for counter = 1 to (maxpages/5)
				pgenumber = (counter*5)
				pgeselect = ""
				if pgenumber = mypage then pgeselect = " selected"
				Response.Write("<option value=""" & scriptname & "?whichpage=" & pgenumber & strQS & """" & pgeselect & ">" & pgenumber & vbCrLf)
			next
			if (maxpages mod 5) > 0 then
				pgeselect = ""
				if maxpages = mypage then pgeselect = " selected"
				Response.Write("<option value=""" & scriptname & "?whichpage=" & maxpages & strQS & """" & pgeselect & ">" & maxpages & vbCrLf)
			end if
			Response.Write("</select>")
			Response.Write("</td></tr></table>" & vbCrLf)
			Response.Write("</form>")
			Response.Write("</td><td nowrap>&#160;&#160;</td>")
		end if
		
		dim pgelow, pgehigh, pgediv

		if maxpages > intPagingLinks then
			pgediv = Int(Abs(intPagingLinks/2))
			pgelow = mypage - pgediv
			pgehigh = mypage + (intPagingLinks - (pgediv + 1))
			if pgelow < 1 then
				pgelow = 1
				pgehigh = pgelow + (intPagingLinks - 1)
			end if
			if pgehigh > maxpages then
				pgehigh = maxpages
				pgelow = pgehigh - (intPagingLinks - 1)
			end if
		else
			pgelow = 1
			pgehigh = maxpages
		end if

		Response.Write("<td nowrap>&#160;")
		if pgelow > 1 then
			response.write("<a href=""" & scriptname & "?whichpage=1" & strQS & """>&lt;&lt;</a>&#160;")
		else
			response.write("&#160;&#160;&#160;&#160;")
		end if
		Response.Write("</td><td>&#160;")
		for counter = pgelow to pgehigh
			if counter <> mypage then
				response.write("&#160;<a href=""" & scriptname & "?whichpage=" & counter & strQS & """>" & counter & "</a>")
			else
				response.write("&#160;" & counter)
			end if
			if counter < pgehigh then response.write("&#160;&#160;|&#160;")
		next
		Response.Write("</td><td nowrap>&#160;")
		if pgehigh < maxpages then
			response.write("&#160;<a href=""" & scriptname & "?whichpage=" & maxpages & strQS & """>&gt;&gt;</a>&#160;")
		else
			response.write("&#160;&#160;&#160;&#160;")
		end if
		Response.Write("</td><td nowrap>&#160;")
		
		' Previous Page Link
		if mypage = 1 then
			response.write(txtPrevious)
		else
			response.write("<a href=""" & scriptname & "?whichpage=" & (mypage - 1) & strQS & """>" & txtPrevious & "</a>")
		end if
		response.write("&#160;|&#160;")
		
		' Next Page Link
		if mypage = maxpages then
			response.write(txtNext)
		else
			response.write("<a href=""" & scriptname & "?whichpage=" & (mypage + 1) & strQS & """>" & txtNext & "</a>")
		end if
		response.write("&#160;|&#160;")
		
		' Reload Page Link
		response.write("<a href=""" & scriptname & "?whichpage=" & mypage & strQS & """>" & txtReload & "</a>")
		Response.Write("</td></tr></table>")


	else
		response.write("<div style=""font-size:6px;"">&#160;</div>")
	end if

end sub

sub getMiniProfile() %>
  <table width="100" align="center">
	<tr>
		<td width="100%" valign="top" align="center" colspan="2" nowrap>
		<span class="fSmall">
		<% If trim(rsd("M_FIRSTNAME")) <> "" or trim(rsd("M_LASTNAME")) <> "" Then 
				response.Write(rsd("M_FIRSTNAME") & " " & rsd("M_LASTNAME") & "<BR>")
			  End If 
				response.Write(txtStatus & ": " & chkIsOnline(rsd("M_NAME"),1) & "<BR>")
				response.Write(txtJoin & ": " & split(strtodate(rsd("M_DATE"))," ")(0) & "<BR>")
			  If trim(rsd("M_CITY")) <> "" Then 
				response.Write(txtCity & ": " & rsd("M_CITY") & "<BR>")
			  end if
			  If trim(rsd("M_STATE")) <> "" Then 

				response.Write(txtState & ": " & rsd("M_STATE") & "<BR>")
			  end If
			  If Trim(rsd("M_COUNTRY")) <> "" Then 
			       Response.Write(rsd("M_COUNTRY") & "<BR>")
			  end If
		%>
		</span>
		</td>
	</tr>
	<tr>
		<td width="50%" align="right" nowrap>
<%			hasIM = "" %>
		<a href="cp_main.asp?cmd=8&member=<% =rsd("MEMBER_ID") %>"> <small>Bio&nbsp;</small><img src="images/icons/icon_profile.gif" height="15" width="15" title="<%= txtViewProf %>" alt="<%= txtViewProf %>" border="0" align="absmiddle"></a>&nbsp;</td><td width="50%" align="left" nowrap><% if chkApp("PM","USERS") and rsd("M_PMSTATUS") = 1 and rsd("M_PMRECEIVE") then %>
		&nbsp;<a href="Javascript:;" onclick="Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=<%= getmemberid(rsd("M_NAME")) %>');"><img src="images/icons/pm.gif" height=17 width=11 title="<%= replace(txtSndPvtMsg,"[%member%]",rsd("M_NAME")) %>" alt="<%= replace(txtSndPvtMsg,"[%member%]",rsd("M_NAME")) %>" border="0" align="absmiddle"><small>&nbsp;<%= txtPM %></small></a><% else %>&nbsp;<% end if %>
		</td></tr>
	<tr><td width="50%" align="right" nowrap>
<%			hasIM = "1" %>
<%	if (lcase(strEmail) = "1" and rsd("M_HIDE_EMAIL") = 0) then 
			if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then  %>
				<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rsd("MEMBER_ID") %>')"><small><%= txtEmail %>&nbsp;</small><img src="images/icons/icon_email.gif" height=15 width=15 title="<%= txtEmlMbr %>" alt="<%= txtEmlMbr %>" border="0" align="absmiddle"></a>&nbsp;
<%			hasIM = "1" %>
<%		end if
		else %>
			&nbsp;<img src="images/spacer.gif" height=15 width=15 alt="" border="0" align="absmiddle">&nbsp;
<%			hasIM = "1" %>
<%	end if %>  
		</td><td width="50%" nowrap align="left">
<%			if strHomepage = "1" then %>
<%				if len(rsd("M_Homepage")) > 7 then %>
        &nbsp;<a href="<% =ChkString(rsd("M_Homepage"),"displayimage") %>" target="_blank"><img src="images/icons/icon_homepage.gif" height=15 width=15 alt="<%= replace(txtVisitHmPg,"[%member%]",rsd("M_NAME")) %>" title="<%= replace(txtVisitHmPg,"[%member%]",rsd("M_NAME")) %>" border="0" align="absmiddle"><small>&nbsp;Web</small></a>
<%			hasIM = "1" %>
<%				end if %>
<%			end if %></td></tr>
	<tr><td width="50%" align="right" nowrap>
<%			if (strMSN = "1") then %>
<%				if Trim(rsd("M_MSN")) <> "" then %>
        <a href="JavaScript:;"><small>msn&nbsp;</small><img src="images/icons/icon_msn.gif" alt="" border="0" align="absmiddle" onClick="openWindow('pop_portal.asp?cmd=7&mode=3&msn=<% =ChkString(replace(rsd("M_MSN"),"@","[no-spam]@"), "displayimage") %>&M_NAME=<% =ChkString(rsd("M_NAME"), "JSurlpath") %>')"></a>&nbsp;
<%			hasIM = "1" %>
<%				end if %>
<%			end if %>
		</td><td width="50%" align="left" nowrap>
<%			if (strAIM = "1") then %>
<%				if Trim(rsd("M_AIM")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=2&AIM=<% =ChkString(rsd("M_AIM"), "JSurlpath") %>&M_NAME=<% =ChkString(rsd("M_NAME"),"urlpath") %>')"><img src="images/icons/icon_aim.gif" height="15" width="15" alt="" border="0" align="absmiddle"><small>&nbsp;aim</small></a>
<%			hasIM = "1" %>
<%				end if %>
<%			end if %></td></tr>
	<tr><td width="50%" align="right" nowrap>
<%			if strICQ = "1" then %>
<%			  if Trim(rsd("M_ICQ")) <> "" then %>
        <a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=1&ICQ=<%= cLng(rsd("M_ICQ")) %>&M_NAME=<% =ChkString(rsd("M_NAME"),"JSurlpath") %>')"><small>icq&nbsp;</small><img src="http://web.icq.com/whitepages/online?icq=<% = ChkString(rsd("M_ICQ"),"display")  %>&img=5" alt="ICQ number" title="ICQ number" border="0" align="absmiddle"></a>&nbsp;
<%			hasIM = "1" %>
<%			  end if %>
<%			end if %>
		</td><td width="50%" align="left" nowrap>
<%			if strYAHOO = "1" then %>
<%			  if Trim(rsd("M_YAHOO")) <> "" then 
					if instr(rsd("M_YAHOO"),"@") then
					Yhoo = ChkString(replace(rsd("M_YAHOO"),"@","[no-spam]@"), "display") 
					else
					Yhoo = ChkString(rsd("M_YAHOO"), "display")
					end if %>
        &nbsp;<a href="http://edit.yahoo.com/config/send_webmesg?.target=<% =ChkString(rsd("M_YAHOO"), "JSurlpath") %>&.src=pg" target="_blank"><img src="images/icons/icon_yahoo.gif" height=15 width=15 alt=""  border="0" align="absmiddle"><small>ahoo</small></a>
<%			hasIM = "1" %>
<%			  end if %>
<%			end if %>
		</td>
	</tr>
</table>
<%
end sub
%>