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

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site announcements
' :::::::::::::::::::::::::::::::::::::::::::::::
function announce_fp()
	intFirst = 1
	strSql = "SELECT * FROM " & strTablePrefix & "ANNOUNCEMENTS WHERE A_START_DATE <= '" & strCurDateString & "' and A_END_DATE >= '" & strCurDateString & "' ORDER BY A_ID DESC;"
	set rsAnn =  my_Conn.Execute(strSql)
	if not rsAnn.EOF then 
	  spThemeTitle=txtAnnouncements
	  spThemeBlock1_open(intSkin)%>
	  <table id="annCat" border="0" cellspacing="1" cellpadding="0" width="100%"><%
  	  Do until rsAnn.EOF
		 A_ID = rsAnn("A_ID")
		 A_SUBJECT = trim(replace(rsAnn("A_SUBJECT"),"''","'"))
		 A_DATE = ChkDate(rsAnn("A_START_DATE"))
		 A_MESSAGE = trim(replace(rsAnn("A_MESSAGE"),"''","'"))
		 A_MESSAGE	= replace(A_MESSAGE,"</p><p>","<br /><br />")
		 A_MESSAGE	= replace(A_MESSAGE,"<p>","")
		 A_MESSAGE	= replace(A_MESSAGE,"</p>","")
		 
		 if intFirst = 1 then
			catHide = ""
			catImg = "min"
			catAlt = txtCollapse
		 else
			catHide = "none"
			catImg = "max"
			catAlt = txtExpand
		 end if %>			
	        <tr>
			<td width="80%" height="25" valign="top" class="tSubTitle"><% If hasAccess(1) Then %><a href="admin_announce.asp?cmd=1&a_id=<%= A_ID %>"><img src="images/icons/icon_edit_topic.gif" align="right" border="0" alt="<%= txtEdit %>" title="<%= txtEdit %>"></a><% End If %><img name="annCat<%=A_ID%>Img" id="annCat<%=A_ID%>Img" src="Themes/<%=strTheme%>/icon_<%=catImg%>.gif" onclick="javascript:mwpHS('annCat','<%=A_ID%>','tbody');" style="cursor:pointer;" title="<%=catAlt%>" alt="<%=catAlt%>" />&nbsp;<b><%=A_SUBJECT%></b>&nbsp;
			</td>
			<td width="20%" valign="middle" class="tSubTitle" align="right"><b><%= A_DATE %></b>
			</td>
			</tr>
			<tbody id="annCat<%=A_ID%>" style="display:<%=catHide%>;">
			<tr>
			<td width="100%" valign="top" colspan="2">
			<div align="justify" class="tPlain"><%= A_MESSAGE %></div>
			</td>
			</tr>
			</tbody>
		<%
		intFirst = 0
		rsAnn.MoveNext
	  Loop 
	  set rsAnn = nothing %>
	</table>
 	<% spThemeBlock1_close(intSkin)
	End if
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site welcome message
' :::::::::::::::::::::::::::::::::::::::::::::::
function welcome_fp()
	if mLev > 0 then
	  w_id = 1
	else
	  w_id = 2
	end if
	strSql = "SELECT * FROM " & strTablePrefix & "WELCOME WHERE W_ID = " & w_id & " AND W_ACTIVE=1"
	set rsWelcome =  my_Conn.Execute (strSql)
    if not rsWelcome.EOF then
	  W_ID = rsWelcome("W_ID")
	  W_TITLE = trim(replace(rsWelcome("W_TITLE"),"''","'"))
	  W_SUBJECT = trim(replace(rsWelcome("W_SUBJECT"),"''","'"))
	  W_SUBJECT	= replace(W_SUBJECT,"[%member%]",strdbntusername)
	  W_MESSAGE = trim(replace(rsWelcome("W_MESSAGE"),"''","'"))
	  W_MESSAGE	= replace(W_MESSAGE,"</p><p>","<br /><br />")
	  W_MESSAGE	= replace(W_MESSAGE,"<p>","")
	  W_MESSAGE	= replace(W_MESSAGE,"</p>","")
	  W_MESSAGE	= replace(W_MESSAGE,"[%member%]",strdbntusername)

	  spThemeMM = "welcom"
	  'spThemeTitle = txtWelcomeTo & " " & strSiteTitle
	  spThemeTitle = W_TITLE
	  spThemeBlock1_open(intSkin) %>
		<div class="tPlain"><%= W_SUBJECT %></div>
		<div class="tPlain"><%= W_MESSAGE %></div>
	  <%
	  spThemeBlock1_close(intSkin)
	End if 
    set rsWelcome = nothing
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site skin changer box
' :::::::::::::::::::::::::::::::::::::::::::::::
function theme_changer()

ssSQL = "select C_STRAUTHOR from portal_colors where C_STRFOLDER = '" & strTheme & "'"
set thmAuth = my_Conn.execute(ssSQL)
if thmAuth.eof then
	strAuth = "anonymous"
else
	strAuth = thmAuth(0)
end if
set thmAuth = nothing

spThemeMM = "sknchgr"
spThemeTitle = txtSknChgr
'spThemeTitle = spThemeTitle & " [" & intSkin & "]"
spThemeBlock1_open(intSkin)
%>
<table><tr>
        <td class="spThemeChanger" align="center" valign="middle">
          <form name="themechanger" method="post" action="<%= Request.ServerVariables("URL") %>">
		  Select Skin:<br /><br />
		   <% 'DoDropDown(fTableName, fDisplayField, fValueField, fSelectValue, fSelectName, fFirstOption)
		    DoSubmitDropDown "PORTAL_COLORS", "C_TEMPLATE", "C_STRFOLDER", "" & strTheme & "", "thm","", "C_SKINLEVEL <= " & mLev & "", "C_TEMPLATE"
			%>
          </form><span class="fSmall"><br /><%= txtAuthor %>:<b> <%= strAuth %> </b></span>
<!-- Skin Levels by wingflap: admin@wingflap.com / http://www.wingflap.com - http://www.planetloser.com copyright 2006 -->
<% if mLev = 0 then
	if getSkinCountDiffByLevel() > 0 then%>
	  <br /><b><a href="policy.asp"><span class="fSmall"><%= txtRegister %></span></a>
	  <span class="fSmall"> for more skins</b></span>
 <% end if
   end if %>
  </td>
</tr></table>
<% spThemeFooter = ""
spThemeBlock1_close(intSkin)
end function 

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site searchbox
' :::::::::::::::::::::::::::::::::::::::::::::::
function search_fp()
%>
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
function RefreshS() {
if (document.SearchForm.news.checked) {
	window.location ="forum_search.asp?mode=news";
} else {
	window.location ="search.asp";
}
}
function checklength() {
if (document.srcform1.search.value.length < 3) {
alert('<%= txtSrchLen %>');
return false;
}
}
// done hiding -->
</script>

<%
spThemeMM = "ssrch"
spThemeTitle= txtSearch
'spThemeTitle = spThemeTitle & " [" & intSkin & "]"
spThemeBlock1_open(intSkin)
  spThemeTitle= txtSrchFor & ":"
  spThemeBlock3_open() %>
<form name="srcform1" action="site_search.asp" method="post" id="srcform1" onsubmit="return checklength()">
    <div class="tPlain" style="text-align:center;"><br />
	<input type="text" name="search" size="15" style="margin-top:5px;" /><br /><br />
      <input type="submit" value=" <%= txtSearch %> " id="searchA" name="searchA" class="button" /><br /><br /></div>
</form>
<%
  spThemeBlock3_close()
spThemeBlock1_close(intSkin)
end function

sub affiliateBanners()
  showHowMany = 10 'ORDER BY ID DESC
  sSQL = "Select * FROM PORTAL_BANNERS WHERE B_LOCATION=2 AND B_ACTIVE=1"
  executeThis(sSQL)
  set rsAB = my_Conn.execute(sSQL)
  if not rsAB.eof then
	spThemeMM = "aff_sm"
    spThemeTitle = txtAffiliates
    spThemeBlock1_open(intSkin)
     %><div>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<% do until rsAB.eof
		strImage = rsAB("B_IMAGE")
		strHover = rsAB("B_ACRONYM")
		intID = rsAB("ID") %>
			<tr><td align="center" height="40"><a target="_blank" title="<%= strHover %>" href="banner_link.asp?id=<%= intID %>"><% If right(strImage,4) = ".swf" Then writeFlash2 strImage,intID,strHover Else response.write("<img alt=""" & strHover & """ name=""abImage"" border=""0"" src=""" & strImage & """ />") end if %></a></td></tr>
	<%     rsAB.movenext
	   loop %>
		</table></div>
<%
    spThemeBlock1_close(intSkin)
  end if
  set rsAB = nothing
end sub

sub login_box() %> 
<div id="login_form" style="display:block;">
<%
if mLev = 0 then
spThemeTitle= txtLogin
spThemeBlock1_open(intSkin) %>
<table border="0" cellpadding="0" cellspacing="0"><tr><td>
<form action="<% =Request.ServerVariables("URL") %>" method="post" id="logmex" name="logmex">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
<td align="center"><br>
  <%= txtUsrName %>: <input class="textbox" type="text" name="Name" size="15" maxlength="25" value="" style="margin:0px;" /><br /><br>
  <%= txtPass %>: <input class="textbox" type="password" name="Password" size="15" maxlength="25" style="margin:0px;" />
<% If SecImage >1 Then %>
<br></td></tr><tr>
<td align="center" height="50">
<img src="includes/securelog/image.asp" alt="<%= txtSecImg %>" title="<%= txtSecImg %>" /><br />
</td></tr><tr>
<td align="center">
<input class="textbox" type="text" name="secCode" size="15" maxlength="8" value="<%= txtSecCode %>" onfocus="javascript:this.value='';" />
<%end if %><br />
</td></tr><tr>
<td align="center"><br>
<input type="checkbox" name="SavePassWord" value="true" checked="checked" />&nbsp;&nbsp;<%= txtSvPass %>
<br /><br>
<input type="submit" value="<%= txtLogin %>" id="logmein" name="logmein" class="btnLogin" /><input type="hidden" name="Method_Type" value="login" /><br />
<%if (lcase(strEmail) = "1") then %>
<br /><a href="password.asp"><%= txtForgotPass %>?<br /><span class="fSmall">Click Here</span></a><br />
<% end if %>
<%if strNewReg = 1 then %>
<br><%= txtNotMember %>?<br><a href="policy.asp"><span class="fAlert"><%= txtRegNow %>!</span></a><br />
<% End If %>
</td></tr>
</table></form></td></tr></table>
<%
spThemeBlock1_close(intSkin)
end if
%></div><%
end sub

%>
