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
<!-- #include file="lang/en/core_admin.asp" --><%
pgType = "manager"
If Session(strCookieURL & "Approval") = "256697926329" Then
'If 12 = 12 Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<!--#INCLUDE file="includes/inc_admin_functions.asp" -->
<%
dim iPgType, iPgText, bgcolor, fpSQL, iMode, strMsg
dim fp_id, fp_name, fp_iname, fp_active, fp_desc, fp_groups, fp_function, fp_column, fp_sticky

iPgType = 3
iMode = 0
strMsg = ""
if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
end if

if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("fp_id") <> "" or  Request("fp_id") <> " " then
	if IsNumeric(Request("fp_id")) = True then
		fp_id = cLng(Request("fp_id"))
	else
		closeAndGo("default.asp")
	end if
end if

if iMode = 2 or iMode = 3 then
  ex = false
  fp_name = chkString(request.Form("fp_name"),"sqlstring")
  fp_iname = chkString(request.Form("fp_iname"),"sqlstring")
  fp_active = chkString(request.Form("fp_active"),"sqlstring")
  fp_desc = chkString(request.Form("fp_desc"),"sqlstring")
  fp_function = chkString(request.Form("fp_funct"),"sqlstring")
  fp_groups = chkString(request.Form("fp_groups"),"sqlstring")
  fp_column = chkString(request.Form("fp_column"),"sqlstring")
  if fp_iname = "" then fp_iname = fp_name
  if fp_name = "" or fp_desc = "" or fp_function = "" then
    'ex = true
	if iMode = 2 then
	  iPgType = 2
	else
	  iPgType = 4
	end if
	'iPgType = 2
	iMode = 0
	strFPmsg = txtCFP02
	strFPmsg = strFPmsg & "<br>" & txtCFP03
  end if
end if

Select case iMode
  case 1 'delete block
	' need to delete from fp_users table as well
	set rsF = my_Conn.execute("select fp_function from PORTAL_FP where id=" & fp_id)
	  f_funct = rsF(0)
	set rsF = nothing
    executeThis("delete from portal_fp where id=" & fp_id)
	delFPusers(f_funct)
    closeandgo("admin_config_fp.asp?cmd=1")
	
  case 2 'add block
    sSQL = "insert into portal_fp ("
	sSQL = sSQL & "fp_name,fp_iname,fp_active,fp_desc,fp_function,fp_groups,fp_column"
	sSQL = sSQL & ")values("
	sSQL = sSQL & "'" & fp_name & "','" & fp_iname & "'," & fp_active & ",'" & fp_desc & "','" & fp_function & "','" & fp_groups & "'," & fp_column & ""
	sSQL = sSQL & ");"
	executeThis(sSQL)
    closeandgo("admin_config_fp.asp?cmd=" & fp_active)
	
  case 3 'edit block
    sSQL = "update portal_fp set "
	sSQL = sSQL & "fp_name='" & fp_name & "'"
	sSQL = sSQL & ", fp_iname='" & fp_iname & "'"
	sSQL = sSQL & ", fp_active=" & fp_active
	sSQL = sSQL & ", fp_desc='" & fp_desc & "'"
	sSQL = sSQL & ", fp_function='" & fp_function & "'"
	sSQL = sSQL & ", fp_groups='" & fp_groups & "'"
	sSQL = sSQL & ", fp_column=" & fp_column
	sSQL = sSQL & " where id=" & fp_id
	executeThis(sSQL)
	if fp_active = 0 then
	  delFPusers(fp_function)
	end if
	'closeandgo("stop")
    closeandgo("admin_config_fp.asp?cmd=" & fp_active)
    'response.Write("Edit complete!")
	
  case 4 'edit default home page settings
    left_sticky = request.Form("left_sticky")
    main_sticky = request.Form("main_sticky")
    right_sticky = request.Form("right_sticky")
    left_col = request.Form("left_select")
    main_col = request.Form("main_select")
    right_col = request.Form("right_select")
    sSQL = "UPDATE PORTAL_FP_USERS SET "
    sSQL = sSQL & "fp_leftcol = '" & left_col & "'"
    sSQL = sSQL & ",fp_maincol = '" & main_col & "'"
    sSQL = sSQL & ",fp_rightcol = '" & right_col & "'"
    sSQL = sSQL & ",fp_leftsticky = '" & left_sticky & "'"
    sSQL = sSQL & ",fp_mainsticky = '" & main_sticky & "'"
    sSQL = sSQL & ",fp_rightsticky = '" & right_sticky & "'"
    sSQL = sSQL & "WHERE fp_uid = 0"
    executeThis(sSQL)
    closeandgo("admin_config_fp.asp?cmd=1")
	
  case 5 'reset all users to the default layout
	strSql = "DELETE FROM PORTAL_FP_USERS WHERE fp_uid <> 0"
	executeThis(strSql)
	strFPmsg = txtCFP04
end select

select case iPgType
  case 0 'front page inactive blocks
    iPgText = txtCFP05
  case 1 'Front Page Active Blocks
    iPgText = txtCFP06
  case 2 'Add Home Page Block
    iPgText = txtCFP07
	iMode = 2
  case 3 'Set Default Home Page Layout
    iPgText = txtCFP08
	iMode = 4
  case 4 'Edit Home Page Block
    iPgText = txtCFP09
	iMode = 3
  case else 'Front Page Active Blocks
    iPgType = 1
    iPgText = txtCFP06
end select

function showActive(num)
	if num = 1 then
	    tmpSho = txtYes
	else
	    tmpSho = txtNo
	end if
	  showActive = tmpSho
end function

function chkFunction(sTemp)
  tmpSt = ""
  if instr(sTemp,":") > 0 then
    tmpSt = split(sTemp,":")(0) & "(""" & split(sTemp,":")(1) & """)"
  else
    tmpSt = sTemp
  end if
  chkFunction = tmpSt
end function
%>
<script type="text/JavaScript">
function delBlock(grp,frmID){
  if (confirm("<%= txtCFP10 %> '"+grp+"'?\n<%= txtCannotBeUndn %>")) {
   for (i=0; i<document.forms.length; i++) {
    if (document.forms[i].name == "fm_"+frmID) {
     document.forms[i].submit();
	}
   }
  }
}
function checkfrm(){
 if (document.forms.addedit.fp_name.value == "") {
 alert("<%= txtCFP11 %>");
	document.forms.addedit.fp_name.focus();
 return false;
 }
 if (!CheckName(document.forms.addedit.fp_name.value)) {
 alert("<%= txtCFP12 %>: \\ / : *  \" < > |");
	document.forms.addedit.fp_name.focus();
 return false;
 }
 
 if (document.forms.addedit.fp_funct.value == "") {
 alert("<%= txtCFP13 %>");
	document.forms.addedit.fp_funct.focus();
 return false;
 }
 if (!CheckThis(document.forms.addedit.fp_funct.value)) {
 alert("<%= txtCFP14 %>:  *  \" < > |");
	document.forms.addedit.fp_funct.focus();
 return false;
 }
 
 if (document.forms.addedit.fp_desc.value == "") {
 alert("<%= txtCFP15 %>");
	document.forms.addedit.fp_desc.focus();
 return false;
 }
 document.forms.addedit.submit();
 }

function chkInput(strStr,params) {
var re = new RegExp("\.(" + params.replace(/,/gi,"|").replace(/\s/gi,"") + ")$","i");
    if(!re.test(strStr)) return false;
	else return true;
}
function CheckThis(str) {
	var re;
	re = /[*'"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
function CheckName(str) {
	var re;
	re = /[\\\/:*'?"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
</script>

<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	fpConfigMenu("1")
  	response.Write("<hr />")
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtCFP16 & "|admin_config_fp.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<% 
spThemeTitle = iPgText
spThemeBlock1_open(intSkin)
if strFPmsg <> "" then
    showUpdResult(strFPmsg)
end if

Select case iPgType
  case 0, 1
    showBlocks(iPgType)
  case 2
    editBlock("Add")
  case 3
    editFPlayout()
  case 4
    sSQL = "select * from PORTAL_FP WHERE id=" & fp_id
	set rsFP = my_Conn.execute(sSQL)
	  fp_name = rsFP("fp_name")
	  fp_iname = rsFP("fp_iname")
	  fp_active = rsFP("fp_active")
	  fp_desc = rsFP("fp_desc")
	  fp_function = rsFP("fp_function")
	  fp_groups = rsFP("fp_groups")
	  fp_column = rsFP("fp_column")
	set rsFP = nothing
    editBlock("Edit")
end Select
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#INCLUDE file="inc_footer.asp" -->
<% else %>
<%     Response.Redirect "admin_login.asp?target=admin_config_fp.asp" %>
<% end if 

' here are the page subroutines
sub delFPusers(num)
  dim mmSQL, tmp_Col
  tmp_Col = ""
  mmSQL = "select fp_leftcol, fp_rightcol, fp_maincol, fp_uid from PORTAL_FP_USERS where fp_leftcol like '%" & num & "%' or fp_maincol like '%" & num & "%' or fp_rightcol like '%" & num & "%'"
  set rsFPd = my_Conn.execute(mmSQL)
  if not rsFPd.eof then
    do until rsFPd.eof
      Col1 = rsFPd("fp_leftcol")
	  Col2 = rsFPd("fp_rightcol")
	  Col3 = rsFPd("fp_maincol")
	  membID = rsFPd("fp_uid")
	  if Col1 <> "" then
		tmp_Col = ""
	    if instr(Col1,",") then
		  alCol = split(Col1,",")
		  cnt = ubound(alCol)
		  for al = 0 to cnt
		    if instr(trim(alCol(al)),num) = 0 then
		    'if num <> trim(alCol(al)) then
		      tmp_Col = tmp_Col & alCol(al) & ","
		    end if
		  next
		else
		  if instr(Col1,num) = 0 then
		    tmp_Col = Col1
		  end if
		end if
		if right(tmp_Col,1) = "," then
		   tmp_Col = left(tmp_Col,len(tmp_Col)-1)
		end if
		Col1 = tmp_Col
	  end if
	  if Col2 <> "" then
		tmp_Col = ""
	    if instr(Col2,",") then
		  arCol = split(Col2,",")
		  cnt = ubound(arCol)
		  for ar = 0 to ubound(arCol)
		    if instr(trim(arCol(ar)),num) = 0 then
		    'if num <> trim(alCol(al)) then
		      tmp_Col = tmp_Col & arCol(ar) & ","
		    end if
		  next
		else
		  if instr(Col2,num) = 0 then
		    tmp_Col = Col2
		  end if
		end if
		if right(tmp_Col,1) = "," then
		   tmp_Col = left(tmp_Col,len(tmp_Col)-1)
		end if
		Col2 = tmp_Col
	  end if
	  if Col3 <> "" then
		tmp_Col = ""
	    if instr(Col3,",") then
		  amCol = split(Col3,",")
		  cnt = ubound(amCol)
		  for ar = 0 to ubound(amCol)
		    if instr(trim(amCol(ar)),num) = 0 then
		    'if num <> trim(alCol(al)) then
		      tmp_Col = tmp_Col & amCol(ar) & ","
		    end if
		  next
		else
		  if instr(Col3,num) = 0 then
		    tmp_Col = Col3
		  end if
		end if
		if right(tmp_Col,1) = "," then
		   tmp_Col = left(tmp_Col,len(tmp_Col)-1)
		end if
		Col3 = tmp_Col
	    end if
		sSQL = "UPDATE PORTAL_FP_USERS SET "
		sSQL = sSQL & "fp_leftcol = '" & Col1 & "'"
		sSQL = sSQL & ",fp_maincol = '" & Col3 & "'"
		sSQL = sSQL & ",fp_rightcol = '" & Col2 & "'"
		sSQL = sSQL & " WHERE fp_uid = " & membID
		'response.Write(sSQL & "<br>")
		executeThis(sSQL)
		rsFPd.movenext
	loop
  end if
  set rsFPd = nothing
end sub

sub editFPlayout()
b_desc = ""
l_options = ""
m_options = ""
r_options = ""
l_select = ""
m_select = ""
r_select = ""

mmSQL = "select * from PORTAL_FP where fp_active = 1 order by fp_name, fp_function"
set rsMM = my_Conn.execute(mmSQL)

if not rsMM.eof then
  do until rsMM.eof
    b_desc = b_desc & "block_descr['" & rsMM("fp_name") & ":" & rsMM("fp_function") & "'] = '" & rsMM("fp_desc") & "';" & vbcrlf
	tmpStr = "<option value=""" & rsMM("fp_name") & ":" & rsMM("fp_function") & """>" & rsMM("fp_name") & "</option>" & vbcrlf
	select case rsMM("fp_column")
	  case 1
	    l_options = l_options & tmpStr
	  case 2
	    m_options = m_options & tmpStr
	  case 3
	    r_options = r_options & tmpStr
	  case 4
	    l_options = l_options & tmpStr
	    r_options = r_options & tmpStr
	end select
    rsMM.movenext
  loop
end if
set rsMM = nothing %>
<script type="text/javascript">
var block_descr = new Array();
<%= b_desc %>
</script>
  <%
  ' populate the select boxes with the default config
  mmSQL = "select * from PORTAL_FP_USERS where fp_uid = 0"
  set rsMM = my_Conn.execute(mmSQL)
  
  if rsMM("fp_leftsticky") <> "" then
    l_stick = split(rsMM("fp_leftsticky"),",")
    for fp = 0 to ubound(l_stick)
	  l_sticky = l_sticky & "<option value=""" & l_stick(fp) & """>" & split(l_stick(fp),":")(0) & "</option>" & vbcrlf
    next
  end if
  if rsMM("fp_mainsticky") <> "" then
    m_stick = split(rsMM("fp_mainsticky"),",")
    for fp = 0 to ubound(m_stick)
	  m_sticky = m_sticky & "<option value=""" & m_stick(fp) & """>" & split(m_stick(fp),":")(0) & "</option>" & vbcrlf
    next
  end if
  if rsMM("fp_rightsticky") <> "" then
    r_stick = split(rsMM("fp_rightsticky"),",")
    for fp = 0 to ubound(r_stick)
	  r_sticky = r_sticky & "<option value=""" & r_stick(fp) & """>" & split(r_stick(fp),":")(0) & "</option>" & vbcrlf
    next
  end if
  l_col = split(rsMM("fp_leftcol"),",")
  m_col = split(rsMM("fp_maincol"),",")
  r_col = split(rsMM("fp_rightcol"),",")
  
  for fp = 0 to ubound(l_col)
	l_select = l_select & "<option value=""" & l_col(fp) & """>" & split(l_col(fp),":")(0) & "</option>" & vbcrlf
  next
  for fp = 0 to ubound(m_col)
	m_select = m_select & "<option value=""" & m_col(fp) & """>" & split(m_col(fp),":")(0) & "</option>" & vbcrlf
  next
  for fp = 0 to ubound(r_col)
	r_select = r_select & "<option value=""" & r_col(fp) & """>" & split(r_col(fp),":")(0) & "</option>" & vbcrlf
  next
set rsMM = nothing %>
<div style="text-align:left;margin:10px;">
<%= txtCFP01 %><br /><br /></div>
<form method="post" action="admin_config_fp.asp" onsubmit="return select_options();">
<input type="hidden" name="cmd" value="3" />
<input type="hidden" name="mode" value="4" />
<input type="hidden" name="name" value="" />
<table border="1" align="center">
<tr class="tTitle"><td valign="center" width="33%">
<b><%= txtCFP17 %></b></td>
<td valign="center" width="33%">
<b><%= txtCFP18 %></b></td>
<td valign="center">
<b><%= txtCFP19 %></b></td></tr>
<!--  START sticky items  -->
<tr><td valign="top">
<table><tr><td valign="top"><%= txtCFP20 %>:<br />
<select multiple="multiple" style="text-align:left;" id="left_sticky" name="left_sticky" size="4">
<%= l_sticky %>
</select>
</td><td align="center"><input type="button" class="details1" onclick="move_up_block('left_sticky');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('left_sticky');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('left_sticky', 'right_sticky');" value=" <%= txtRight %> " /><br />
<input type="button" class="details1" onclick="remove_block('left_sticky');" value="<%= txtRemove %>" /><br />
<input type="button" class="details1" onclick="move_left_right_block('left_sticky', 'left_select');" value="<%= txtUnstick %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td valign="top"><%= txtCFP21 %>:<br />
<select multiple="multiple" style="text-align:left;" id="main_sticky" name="main_sticky" size="4">
<%= m_sticky %>
</select>
</td><td align="center"><input type="button" class="details1" onclick="move_up_block('main_sticky');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('main_sticky');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="remove_block('main_sticky');" value="<%= txtRemove %>" /><br />
<input type="button" class="details1" onclick="move_left_right_block('main_sticky', 'main_select');" value="<%= txtUnstick %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td valign="top"><%= txtCFP22 %>:<br />
<select multiple="multiple" style="text-align:left;" id="right_sticky" name="right_sticky" size="4">
<%= r_sticky %>
</select>
</td><td align="center"><input type="button" class="details1" onclick="move_up_block('right_sticky');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('right_sticky');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_sticky', 'left_sticky');" value=" <%= txtLeft %> " /><br />
<input type="button" class="details1" onclick="remove_block('right_sticky');" value="<%= txtRemove %>" /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_sticky', 'right_select');" value="<%= txtUnstick %>" />
</td></tr></table>
</td></tr>
<!--  end sticky items  -->

<tr><td valign="top">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="left_select" name="left_select" size="10">
<%= l_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_left_right_block('left_select', 'left_sticky');" value=" <%= txtSticky %> " /><br />
<input type="button" class="details1" onclick="move_up_block('left_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('left_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('left_select', 'right_select');" value=" <%= txtRight %> " /><br />
<input type="button" class="details1" onclick="remove_block('left_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="main_select" name="main_select" size="10">
<%= m_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_left_right_block('main_select', 'main_sticky');" value=" <%= txtSticky %> " /><br />
<input type="button" class="details1" onclick="move_up_block('main_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('main_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="remove_block('main_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="right_select" name="right_select" size="10">
<%= r_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_left_right_block('right_select', 'right_sticky');" value=" <%= txtSticky %> " /><br />
<input type="button" class="details1" onclick="move_up_block('right_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('right_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_select', 'left_select');" value=" <%= txtLeft %> " /><br />
<input type="button" class="details1" onclick="remove_block('right_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td></tr>
<tr><td>
<select style="text-align:left;" id="left_add" name="left_add" onchange="show_description('left_add');">
<option value=""><%= txtAddLftCol %>...</option>
<%= l_options %>
</select><br>
<input type="button" class="details1" onclick="add_block('left_select', 'left_add');" value="<%= txtAdd %>" />
</td><td>
<select style="text-align:left;" id="main_add" name="main_add" onchange="show_description('main_add');">
<option value=""><%= txtAddMnCol %>...</option>
<%= m_options %>
</select><br>
<input type="button" class="details1" onclick="add_block('main_select', 'main_add');" value="<%= txtAdd %>" />
</td><td>
<select style="text-align:left;" id="right_add" name="right_add" onchange="show_description('right_add');">
<option value=""><%= txtAddRtCol %>...</option>
<%= l_options %>
</select><br>
<input type="button" class="details1" onclick="add_block('right_select', 'right_add');" value="<%= txtAdd %>" />
</td></tr>
<tr><td colspan="3"><div id="instructions"></div>

<center><input type="submit" value="<%= txtCFP23 %>" /></center>
</td></tr>
</table>
</form><br />
<%
end sub

sub editBlock(fMode) 
   iRead = ""
  if fMode = "Edit" then
    iRead = ""
  end if %>
<form name="addedit" id="addedit" method="post" action="admin_config_fp.asp" onSubmit="checkfrm();return false">
  <center><%= strMsg %></center>
  
  <table width="450" border="0" cellspacing="3" cellpadding="0" align="center">
  <tr> 
    <td width="45%" align="right"><%= txtCFP24 %>: </td>
    <td width="45%">
        <input name="fp_name" type="text" id="fp_name" value="<%= fp_name %>">
        <input name="fp_iname" type="hidden" value="<%= fp_iname %>">
      </td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr> 
    <td align="right"><%= txtCFP25 %>: </td>
    <td>
        <input name="fp_funct" type="text" id="fp_funct" value="<%= fp_function %>"<%= iRead %>>
      </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right" nowrap><%= txtCFP26 %>: </td>
    <td>
        <textarea name="fp_desc" cols="30" rows="3" wrap="VIRTUAL" id="fp_desc"><%= fp_desc %></textarea>
      </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right"><%= txtActive %>: </td>
    <td>
        <select name="fp_active" id="fp_active">
		<% If fp_active = 1 Then %>
          <option value="1" selected><%= txtYes %></option>
          <option value="0"><%= txtNo %></option>
		<% Else %>
          <option value="1"><%= txtYes %></option>
          <option value="0" selected><%= txtNo %></option>
		<% End If %>
        </select>
      </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right"><%= txtCFP27 %>: </td>
    <td>
        <select name="fp_column" id="fp_column">
          <option value="1"<% If fp_column = 1 Then response.Write(" selected") %>><%= txtLeft %></option>
          <option value="2"<% If fp_column = 2 Then response.Write(" selected") %>><%= txtMain %></option>
          <option value="3"<% If fp_column = 3 Then response.Write(" selected") %>><%= txtRight %></option>
          <option value="4"<% If fp_column = 4 Then response.Write(" selected") %>><%= txtEitherSide %></option>
        </select>
      </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right">&nbsp;
        <input name="cmd" type="hidden" value="1">
        <input name="mode" type="hidden" value="<%= iMode %>">
        <input name="fp_id" type="hidden" value="<%= fp_id %>">
        <input name="fp_groups" type="hidden" value="3">
	</td>
    <td>
        <input type="submit" name="Submit" value=" <%= fMode & "&nbsp;" & txtCFP28 %> ">
      </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</form><%
end sub

sub showBlocks(typ) %>
<table border="1" cellpadding="3" cellspacing="0" width="100%" align="center" class="tCellAlt1">
<tr class="tCellAlt0">
	<td width="20%"><b><%= txtCFP24 %></b></td>
	<td width="20%"><b><%= txtCFP25 %></b></td>
	<td width="45%"><b><%= txtDesc %></b></td>
	<td width="15%" align="center"><b></b></td>
</tr>
<% 
bgcolor = "tCellAlt2"
fpSQL = "select * from PORTAL_FP where fp_active = " & typ & " order by fp_function"
 set rsFP = my_Conn.execute(fpSQL)
 if not rsFP.eof then
  do until rsFP.eof
    if bgcolor = "tCellAlt2" then
	  bgcolor = "tCellAlt1"
	else
	  bgcolor = "tCellAlt2"
	end if
	response.Write("<form name=""fm_" & rsFP("id") & """ method=""post"" action=""admin_config_fp.asp"">" & vbCrLf)
	response.Write("<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />" & vbCrLf)
	response.Write("<input type=""hidden"" name=""mode"" value=""1"" />" & vbCrLf)
	response.Write("<input type=""hidden"" name=""fp_id"" value=""" & rsFP("id") & """ />" & vbCrLf)
	response.Write("<tr class=""" & bgcolor & """><td>" & rsFP("fp_name") & "</td>" & vbCrLf)
	response.Write("<td>" & chkFunction(rsFP("fp_function")) & "</td>" & vbCrLf)
	response.Write("<td>" & rsFP("fp_desc") & "</td>" & vbCrLf)
	response.Write("<td align=""center""><a href=""admin_config_fp.asp?cmd=4&fp_id=" & rsFP("id") & """>")
	response.Write("<img src=""images/icons/icon_edit_topic.gif"" alt=""" & txtCFP29 & """ border=""0"" title=""" & txtCFP29 & """></a>")
	response.Write("&nbsp;<a href=""javascript:delBlock('" & rsFP("fp_name") & "','" & rsFP("id") & "');"">")
	response.Write("<img src=""images/icons/icon_trashcan.gif"" alt=""" & txtCFP30 & """ border=""0"" title=""" & txtCFP30 & """></a>")
	response.Write("</td></tr></form>" & vbCrLf)
    rsFP.movenext
  loop	
 else
	response.Write("<tr><td width=""100%"" colspan=""4"">" & vbCrLf)
	response.Write("<b>" & txtCFP31 & "</b><td><tr>" & vbCrLf)
 end if
    response.Write("</table>" & vbCrLf)
end sub
%>