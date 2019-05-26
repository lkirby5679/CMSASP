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

strDefaultFontFace = "Ariel"
strDefaultFontSize = "3"
strDefaultFontColor = "#336699"
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<table cellPadding="0" cellSpacing="0" border="0" width="100%" align="center">
<tr><td class="tCellAlt1" width="100%" valign="top"><font size="<% =strDefaultFontSize %>"><b>Website Name:</b>&nbsp;<%= strSiteTitle %></font><br><font size="<% =strDefaultFontSize %>"><b>Version:</b>&nbsp;<%= strWebSiteVersion %></font><br><% If request("mode") = "sd" Then %><br><font size="<% =strDefaultFontSize %>"><b>Downloaded from:</b>&nbsp;<a href="http://www.<%= stMx %><%= stWb %><%= stPl %><%= lcase("net") %>" target="_blank">www.<%= stMx %><%= stWb %><%= stPl %><%= lcase("net") %> v<%= "1" %>.</a></font><br><font size="<% =strDefaultFontSize %>"><b><%= w_DL %> @:</b>&nbsp;<a href="http://www.<%= stMx %><%= stWb %><%= stPl %><%= lcase("net") %>" target="_blank">www.<%= stMx %><%= stWb %><%= stPl %><%= lcase("net") %> v<%= "1" %>.</a></font><br><% End If %></td></tr></table><!-- SkyPortal Version Page End-->
<!--#INCLUDE FILE="inc_footer.asp" -->
