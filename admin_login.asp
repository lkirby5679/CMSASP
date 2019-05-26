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

%><!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/core_admin.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td class="leftPgCol" nowrap>
	<% 
	intSkin = getSkin(intSubSkin,1)
	menu_fp() %>
    </td>
    <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdmLogin & "|admin_login.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  spThemeBlock1_open(intSkin)
	RequestMethod = Request.ServerVariables("Request_method")

IF RequestMethod = "POST" Then
	sName = strDBNTUserName
	Password = ChkString(pEncrypt(pEnPrefix & Request.Form("Password")), "SQLString")
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE M_NAME = '" & trim(sName) & "' AND"
	strSql = strSql & " M_PASSWORD = '" & trim(Password) & "' AND"
	strSql = strSql & " M_LEVEL = 3 AND M_STATUS = 1"
	
	Set dbRs = my_Conn.Execute(strSql)
		
	If not(dbRS.EOF) and ChkQuoteOk(Name) and ChkQuoteOk(Password) Then 
	  If (SecImage < 3) OR (SecImage > 2 and DoSecImage(Ucase(request.form("SecCode"))) = 1) Then
%>
<p align="center"><span class="fTitle">Login was successful!</span></p>
<% Session(strCookieURL & "Approval") = "256697926329"

if trim(Request.form("target")) = "" then
  closeAndGo("admin_home.asp")
else
  closeAndGo(request.form("target"))
end if
%> 
<% 	  Else %>
		<p align="center"><span class="fSubTitle"><%= txtSecCodeBad %></span></p>
<%    end if%>
<% 	Else %>
	  <center>
	  <p align="center"><span class="fSubTitle"><%= txtThereIsProb %></span></p>
	  <p align="center"><span class="fSubTitle"><%= txtNoPermViewPg %></span></p>
	  <p align="center"><%= txtLogErrTryAgn %></p>
	  </center>
<%  End IF
End IF  %>
<script language="JavaScript" type="text/JavaScript">
function focuspass() { document.forms.Form1.Password.focus(); }
window.onload=focuspass;
</script>
<form action="admin_login.asp" method="post" id="Form1" name="Form1">
<table border="0" cellspacing="2" cellpadding="0" align="center">
  <tr>
    <td align="center" colspan="2" class="tTitle"><b><%= txtAdmLogin %></b></td>
  </tr>
  <tr>
	<td  class="tCellAlt1" align="right" nowrap><b><%= txtUsrName %>:</b></td>
	<td class="tCellAlt1"><input type="text" name="Name" size="20" value="<%= strDBNTUserName %>"></td>
  </tr>
  <tr>
	<td align="right" class="tCellAlt1" nowrap><b><%= txtPass %>:</b></td>
	<td class="tCellAlt1"><input type="Password" name="Password" size="20">
	<input type="hidden" name="target" value="<%= chkstring(request.querystring("target"),"clean") %>">
	</td>
</tr>
<% If SecImage > 2 Then %>
  <TR>         
	 <TD align=center colspan="2" > 		
		<img src="includes/securelog/image.asp" />
	 </td>	 	
  </TR> 
  <TR>         
	 <TD align=center colspan="2" > 		
		<%= txtEntrSecImg %>
	 </td>	 	
  </TR> 	 
  <TR>	 	    
	 <TD colspan="2"><input name="secCode" size="20" value="<%= txtSecCode %>" onFocus="javascript:this.value='';"></td>
  </TR>	  
<%End If %> 
  <tr>
    <td colspan="2" align="right"><input type="submit" value="<%= txtLogin %>" id="Submit1" name="Submit1" class="button"></td>
  </tr>
</table>
</form>
<% spThemeBlock1_close(intSkin) %>
    </td>
  </tr>
</table>

<!--#INCLUDE file="inc_footer.asp" -->