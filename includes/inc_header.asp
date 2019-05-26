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
'strHeaderType = 2
if strHeaderType = 2 or strHeaderType = 1 then
	' Select random banner to display on page first
	strSQL = "SELECT * FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.B_ACTIVE=1 and " & strTablePrefix & "BANNERS.B_LOCATION=1"
	set rsLinks = server.CreateObject("adodb.recordset")
	rsLinks.CursorType = 3
	rsLinks.CursorLocation = 3
	rsLinks.LockType = 3
	rsLinks.Open strSQL, my_Conn
	if not rsLinks.eof then
		numLinks = rsLinks.recordcount
		if numLinks > 1 then
			Randomize
			RndNumber = Int(Rnd * numLinks) 
			rsLinks.move RndNumber
		end if
		bannerTxt = chkstring(rsLinks("B_ACRONYM"),"display")
		bannerID = rsLinks("ID")
		bannerImg = chkstring(rsLinks("B_IMAGE"),"displayimage")
		bannerName = chkstring(rsLinks("B_NAME"),"display")
		activeBanners = true
'		rsLinks.close
'		set rsLinks = nothing
	else
		activeBanners = false
	end if
	if strHeaderType = 2 and activeBanners then 'add 1 to impression count
		sSql = "UPDATE " & strTablePrefix & "BANNERS SET " & strTablePrefix & "BANNERS.B_IMPRESSIONS = " & strTablePrefix & "BANNERS.B_IMPRESSIONS + 1  WHERE " & strTablePrefix & "BANNERS.ID=" &  bannerID
		on error resume next
		my_Conn.execute(sSql)
		on error goto 0
	end if
end if

if strheaderType = 1 and activeBanners = true then  'random rotating banners
	strSQL = "SELECT B_IMAGE, B_ACRONYM, ID FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.B_ACTIVE=1 and " & strTablePrefix & "BANNERS.B_LOCATION=1 and " & strTablePrefix & "BANNERS.B_IMAGE NOT LIKE '%.swf'"
	set rsLinks = my_Conn.execute(strSQL)
	if not rsLinks.eof then
		Response.Write("<script type=""text/javascript"">") & vbcrlf
		Response.Write("<!-- Begin") & vbcrlf
		Response.Write("var interval = 10; // delay between rotating images (in seconds)") & vbcrlf
		Response.Write("var random_display = 1; // 0 = no, 1 = yes") & vbcrlf
		Response.Write("interval *= 1000;") & vbcrlf
		Response.Write("var image_index = 0;") & vbcrlf
		Response.Write("image_list = new Array();") & vbcrlf
		Response.Write("link_list = new Array();") & vbcrlf
		Response.Write("text_list = new Array();") & vbcrlf
		Response.Write("var url = ""banner_link.asp?id=2"";") & vbcrlf
		Do until rsLinks.eof
			Response.Write("image_list[image_index++] = new imageItem("""&rsLinks("B_IMAGE")&""");") & vbcrlf
			Response.Write("text_list[image_index] = """&rsLinks("B_ACRONYM")&""";")  & vbcrlf
			Response.Write("link_list[image_index] = "&rsLinks("ID")&";") & vbcrlf
			rsLinks.movenext
		loop
		Response.Write("var number_of_image = image_list.length;") & vbcrlf
		Response.Write("function imageItem(image_location) {") & vbcrlf
		Response.Write("this.image_item = new Image();") & vbcrlf
		Response.Write("this.image_item.src = image_location;") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("function get_ImageItemLocation(imageObj) {") & vbcrlf
		Response.Write("return(imageObj.image_item.src);") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("function generate(x, y) {") & vbcrlf
		Response.Write("var range = y - x + 1;") & vbcrlf
		Response.Write("return Math.floor(Math.random() * range) + x;") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("function getNextImage() {") & vbcrlf
		Response.Write("if (random_display) {") & vbcrlf
		Response.Write("image_index = generate(0, number_of_image-1);") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("else {") & vbcrlf
		Response.Write("image_index = (image_index+1) % number_of_image;") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("var new_image = get_ImageItemLocation(image_list[image_index]);") & vbcrlf
		Response.Write("return(new_image);") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("function rotateImage(place) {") & vbcrlf
		Response.Write("var new_image = getNextImage();") & vbcrlf
		Response.Write("document[place].src = new_image;") & vbcrlf
		Response.Write("url = ""banner_link.asp?id=""+link_list[image_index+1];+""""") & vbcrlf
		Response.Write("document[place].alt = """"+text_list[image_index+1]+"""";") & vbcrlf
		Response.Write("var recur_call = ""rotateImage('""+place+""')"";") & vbcrlf
		Response.Write("setTimeout(recur_call, interval);") & vbcrlf
		Response.Write("}") & vbcrlf
		Response.Write("//  End -->") & vbcrlf
		Response.Write("</script>") & vbcrlf
		'Response.Write("</head>") & vbcrlf
		if mlev > 0 Then 'they are logged in. rotate the banners
		  spThemeBodyTag = spThemeBodyTag & " onLoad=""rotateImage('bImage')"""
		else 'they are not logged in or they are a guest
		  if strLoginType <> 0 then ' login box is not in the header, rotate the banners
		    spThemeBodyTag = spThemeBodyTag & " onLoad=""rotateImage('bImage')"""
		  end if
		end if
	end If
end if %>
	</head>
	<body<%=spThemeBodyTag%>>
<%
if strHeaderType = 1 or strHeaderType = 2 then
	rsLinks.Close
	Set rsLinks = Nothing
end if %>
<a name="top"></a>
<% spThemeStart() %>
<% headerTop() %>
<% spThemeHeader_open() %>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
	<td align="left" valign="middle"><a href="default.asp"><img title="<% =strSiteTitle %>" alt="<% =strSiteTitle %>" border="0" src="<%= strHomeUrl %>Themes/<%= strTheme %>/<%= subTheme %><%= strTitleImage %>" /></a></td>
    <td align="right" valign="top">
	  	<% 
		if strLoginType = 0 and strdbntusername = "" and strAuthType <> "nt" and strNewReg = 1 then
			showloginbox()
		else			
		  Select Case strHeaderType
			case 0
				shoNotta()
			case 1, 2
				showheaderBanner()
			case 3
				showIcons()
			case 4
				showOther()
			case else
				shoNotta()
		  End Select
		end if
		 %>
	</td>
  </tr>
</table>
<% spThemeHeader_close() %>

<% sub shoNotta() %>
<table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
<tr><td width="100%">&nbsp;</td></tr></table>
<% end sub %>

<% sub showHeaderBanner() %>
		<% If activeBanners = true Then %><div class="sp_Banner"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
			<tr><td align="center" height="50"><% If strHeaderType = 1 then %><a href="#" onclick="window.open(url,'BannerWin');" name="banner"><img alt="" name="bImage" border="0" src="<%= bannerImg %>" /></a><% End If %><% If strHeaderType = 2 Then %><a target="_blank" title="<%= bannerTxt %>" href="banner_link.asp?id=<%= bannerID %>"><% If right(bannerImg,4) = ".swf" Then writeFlash(bannerImg) Else response.write("<img alt="""" name=""bImage"" border=""0"" src=""" & bannerImg & """ />") end if %></a><% End If %></td></tr></table></div><% End If %>
<% end sub %>

<% Sub writeFlash(swfImg) %>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="468" height="60" id="Flash_Banner" align=""><param name=movie value="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bannerID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bannerName) %>"><param name=quality value=high><embed src="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bannerID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bannerName) %>" quality="high" name="Flash_Banner" height="60" width="468" pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></object>
<% end sub %>

<% Sub showloginbox() %>
    <table class="sp_Header_Login" cellpadding="0" cellspacing="0" style="border-collapse: collapse;" align="right">
		<tr>
		  <td align="right" valign="middle">
          <table width="100%" border="0" cellpadding="3" cellspacing="0">
            <form action="<% =Request.ServerVariables("URL") %>" method="post" id="formb1" name="formb1">
              <input type="hidden" name="Method_Type" value="login" />
              <tr> 
                <td width="90" align="center" valign="middle" class="tCellAlt2"><b>&nbsp;<%= txtUsrName %>:</b><br>
                  &nbsp;<input class="textbox" type="text" name="Name" size="10" />
                </td>
                <td width="90" align="center" valign="middle" class="tCellAlt2"><b><%= txtPass %>:</b><br>
                  <input class="textbox" type="password" name="Password" size="10" />
                </td>
                <td width="75" align="center" valign="middle" class="tCellAlt2">&nbsp;<% if strGfxButtons <> 0 then %><input src="images/clear.gif" class="imgLogin" type="image" value="<%= txtLogin %>" id="submitx1" name="submitx1" border="0" hspace="1" /><% else if strGfxButtons = "0" then %><input class="btnLogin" type="submit" value="<%= txtLogin %>" id="submitx1" name="submitx1" /><%	end if end if %></td>
              </tr>
              <tr> 
                <td colspan="3" align="center" class="tCellAlt2"> 
                  <input type="checkbox" name="SavePassWord" value="true" checked />
                  <span class="fSmall"><%= txtSvPass %>&nbsp;&nbsp;</span>
                  <%if (lcase(strEmail) = "1") then %>
                  <a href="password.asp"><span class="fSmall"><%= txtForgotPass %>?</span></a>&nbsp;&nbsp; 
                  <% end if 
				  if strNewReg = 1 then %>
                  <br><span class="fSmall"><%= txtNotMember %>?</span> 
                  <a href="policy.asp"><span class="fSmall"><%= txtRegNow %>!</span></a>
				  <% End If %>
				  </td>
              </tr>
            </form>
          </table>
                </td>
			</tr>
    </table>
<% End Sub %>

<% sub showIcons() %>
<table width="84%" border="0" cellspacing="0" cellpadding="0" class="sp_headerIcons" align="center">
  <tr align="center" valign="middle"> 
    <td width="12%"><a href="fhome.asp" title="<%= txtView %> <%= txtForum %>"><img title="<%= txtView %> <%= txtForum %>" alt="<%= txtView %> <%= txtForum %>" src="Themes/<%= strTheme %>/forums.gif" border="0" /></a></td>
    <td width="12%"><a href="events.asp" title="<%= txtView %> <%= txtCalendar %>"><img title="<%= txtView %> <%= txtCalendar %>" alt="<%= txtView %> <%= txtCalendar %>" src="Themes/<%= strTheme %>/events.gif" border="0" /></a></td>
    <td width="12%"><a href="article.asp" title="<%= txtView %> Articles"><img title="<%= txtView %> Articles" alt="<%= txtView %> <%= txtArticles %>" src="Themes/<%= strTheme %>/articles.gif" border="0" /></a></td>
    <td width="12%"><a href="dl.asp" title="<%= txtView %> <%= txtDownloads %>"><img title="<%= txtView %> <%= txtDownloads %>" alt="<%= txtView %> <%= txtDownloads %>" src="Themes/<%= strTheme %>/dl.gif" border="0" /></a></td>
    <td width="12%"><a href="links.asp" title="<%= txtView %> <%= txtLinks %>"><img title="<%= txtView %> <%= txtLinks %>" alt="<%= txtView %> <%= txtLinks %>" src="Themes/<%= strTheme %>/links.gif" border="0" /></a></td>
    <td width="12%"><a href="pic.asp" title="<%= txtView %> <%= txtPics %>"><img title="<%= txtView %> <%= txtPics %>" alt="<%= txtView %> <%= txtPics %>" src="Themes/<%= strTheme %>/pic.gif" border="0" /></a></td>
    <td width="12%"><a href="classified.asp" title="<%= txtView %> <%= txtClassifieds %>"><img title="<%= txtView %> <%= txtClassifieds %>" alt="<%= txtView %> <%= txtClassifieds %>" src="Themes/<%= strTheme %>/features.gif" border="0" /></a></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td><a href="fhome.asp" title="<%= txtView %> <%= txtForum %>"><b><%= txtForum %></b></a></td>
    <td><a href="events.asp" title="<%= txtView %> <%= txtCalendar %>"><b><%= txtEvents %></b></a></td>
    <td><a href="article.asp" title="<%= txtView %> <%= txtArticles %>"><b><%= txtArticles %></b></a></td>
    <td><a href="dl.asp" title="<%= txtView %> <%= txtDownloads %>"><b><%= txtDownloads %></b></a></td>
    <td><a href="links.asp" title="<%= txtView %> <%= txtLinks %>"><b><%= txtLinks %></b></a></td>
    <td><a href="pic.asp" title="<%= txtView %> <%= txtPics %>"><b><%= txtPics %></b></a></td>
    <td><a href="classified.asp" title="<%= txtView %> <%= txtClassifieds %>"><b><%= txtClassifieds %></b></a></td>
  </tr>
</table>
<% end sub %>

<% sub showOther() %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="100%">&nbsp;
<!-- Add code here to display in the header area
	 when "other" is selected as the Header Type -->
</td></tr></table>
<% end sub %>