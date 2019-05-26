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
curPageType = "core"
%>
<!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<%
PageTitle = txtMemProf
CurPageInfoChk = "1"
function CurPageInfo ()
If Request.QueryString("mode") = "display" Then
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = txtMemProf
	PageAction = txtViewing & "<br>" 
	PageLocation = "pop_profile.asp?" & strOnlineQueryString & ""
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
else
    CurPageInfo = txtEditing & "<br>" & txtMemProf
end if
end function 
if Request.QueryString("mode") = "goEdit" then
  hasEditor = true
  strEditorElements = "Sig"
end if
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<% 

if strDBNTUserName = "" then
	doNotLoggedInForm
else

  if strAuthType = "nt" then
	if ChkAccountReg() <> "1" then %>
	  <p align="center">
	  <b>Note:</b> This NT account has not been registered yet, thus the profile is not available.<br>
	  If this is your account, <a href="register.asp?mode=Register"><b>click here</b></a> to register.</p>		
	  <!--#INCLUDE FILE="inc_footer.asp" -->
	  <% 
	  Response.End 
	end if
  end if

end if
%><!--#INCLUDE FILE="inc_footer.asp" -->