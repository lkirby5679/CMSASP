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
<!--include file="modules/forums/fp_forums.asp" -->
<!--include file="modules/downloads/fp_dl.asp" -->
<!--include file="modules/articles/fp_articles.asp" -->
<!--include file="modules/events/fp_events.asp"-->
<!--include file="modules/pictures/fp_pic.asp"-->
<!--include file="modules/links/fp_links.asp"-->
<!--include file="modules/classifieds/fp_classified.asp"-->
<% 

function cntPendTsks()
  PTcnt = 0
  ' Pending Articles count
  'PTcnt = PTcnt + getCount("ARTICLE_ID","ARTICLE","SHOW=0")
  ' Pending Downloads count
  'PTcnt = PTcnt + getCount("DL_ID","DL","SHOW=0 OR BADLINK<>0")
  ' Pending Pictures count
  'PTcnt = PTcnt + getCount("PIC_ID","pic","SHOW=0 OR BADLINK<>0") 
  ' Pending Classifieds count
  'PTcnt = PTcnt + getCount("CLASSIFIED_ID","CLASSIFIED","SHOW=0 OR BADLINK<>0")
  ' Pending Links count
  'PTcnt = PTcnt + getCount("LINK_ID","LINKS","SHOW=0 OR BADLINK<>0")
  
  cntPendTsks = "&nbsp;(" & PTcnt & ")"
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site projects box
' :::::::::::::::::::::::::::::::::::::::::::::::
function projects_fp()
spThemeMM = "prjct"
spThemeTitle= txtProjStat
spThemeBlock1_open(intSkin)%>
<table border="0" width="100%"><tr><td width="100%" class="tCellAlt1">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td><b><%= strSiteTitle %></b></td></tr>
<tr><td align="left"><table width="100%" border="1" cellspacing="0" cellpadding="0" class="tBorder"><tr><td width="95%" bgcolor="#CCCCCC"><img src="images/icons/bar.gif" width="87%" height="15" alt="" /></td><td bgcolor="whitesmoke"><span class="fSmall">87%</span></td></tr></table></td></tr>
<tr><td><b><%= txtHapBDay %></b></td></tr>
<tr><td align="left"><table width="100%" border="1" cellspacing="0" cellpadding="0" class="tBorder"><tr><td width="95%" bgcolor="#CCCCCC"><img src="images/icons/bar.gif" width="35%" height="15" alt="" /></td><td bgcolor="whitesmoke"><span class="fSmall">35%</span></td></tr></table></td></tr>
</table>
</td></tr></table>
<%spThemeBlock1_close(intSkin)
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		OTHER LINKS box
' :::::::::::::::::::::::::::::::::::::::::::::::
function others_fp()
spThemeMM = "othrs"
spThemeTitle= "Support SkyPortal"
spThemeBlock1_open(intSkin)%>
<table border="0" cellpadding="5" cellspacing="0"><tr>
  <td align="center">
  <p>Please help support the continued development of SkyPortal by making your donation today.</p>
  <p><a href="http://www.skyportal.net/site_donation.asp"><img src="http://www.skyportal.net/images/donation_sp.gif" border="0" alt="" title="Help support the SkyPortal Development" width="172" height="62" /></a></p>
  </td></tr></table>
<%spThemeBlock1_close(intSkin)
end function

function m_aspin()
spThemeTitle = "Rate SkyPortal"
spThemeBlock1_open(intSkin) %>
    <table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr> 
        <td align="left" valign="middle">
		<span style="font-size:8pt;font-family:Arial;">If you use SkyPortal and think we are the best around, and you want everyone to know it, please vote for us at Aspin.com!<br>&nbsp;</span>
		 <table width="100%" align=right border=0 cellpadding=1 cellspacing=0>
		  <tr><td align=center>
		    <table width="100%" border=0 cellpadding=3 cellspacing=0>
			  <tr><td align=center nowrap>
			    <font style="font-size:10pt;font-family:Arial;"><b>Rated:</b> <a href="http://www.Aspin.com/func/review?id=6559210"><img src=http://ratings.Aspin.com/getstars?id=6559210 border=0></a>  <font style="font-size:8pt;"><br>by <a href="http://www.Aspin.com" target="_blank">Aspin.com</a> users<br></font></font> </td></tr>
			  <tr nowrap><form action="http://www.Aspin.com/func/review/write?id=6559210" method="post" target="_blank"><td align=center> <font style="font-size:10pt;font-family:Arial;">What do you think?</font><br> <select name="VoteStars"><option>5 Stars<option>4 Stars<option>3 Stars<option>2 Stars<option>1 Star</select>&nbsp;&nbsp;<input type=submit value="Vote">  </td></form></tr>
			</table>
		  </td></tr>
		 </table> 
        </td>
      </tr>
    </table>
<%
spThemeBlock1_close(intSkin)
end function

Sub writeFlash2(swfImg,bID,bName) %>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="88" height="31" id="abImage" align=""><param name=movie value="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bName) %>"><param name=quality value=high><embed src="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bName) %>" quality="high" name="abImage" height="31" width="88" pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></object>
<% 
end sub

sub modFeatures()
    spThemeBlock1_open(intSkin) %>
	<div style="padding:6px;text-align:left;">All the themeblocks that you see to the left, right, on top (this block) and on bottom of the MAIN page block are controlled by a single <b>'Modules/<%= CurPageTitle %>/<%= CurPageTitle %>_custom.asp'</b> file. You can show any themeblock that is available on the homepage here as well. Or you can create your own function. This way, people can change the layout to what they like, and also keeping their layout in a file that will not be included with future upgrades.<br><br>The subroutine for this block is called <b>modFeatures()</b> and is located in <b>fp_custom.asp</b> at the very bottom of the file. The call for this subroutine is called <b>modFeatures()</b> and is located in <b>'Modules/<%= CurPageTitle %>/<%= CurPageTitle %>_custom.asp'.</b></div>
<%  spThemeBlock1_close(intSkin)
end sub

' insert new functions and subs above this line

%>
