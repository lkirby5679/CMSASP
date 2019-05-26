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

sub displayProfile(iID)
		if strUserMemberID <> iID and strUserMemberID > 0 then
			lastviewid = chkString(Request.Cookies(strCookieURL & "lastviewid"),"sqlstring")
			if lastviewid = "" then lastviewid = 0
			
			if lastviewid <> cLng(Request.QueryString("id")) then
				'update page views
				strSql = "Update " & strMemberTablePrefix & "MEMBERS"
				strSql = strSql & " Set M_PAGE_VIEWS = M_PAGE_VIEWS + 1"
				strSql = strSql & " WHERE MEMBER_ID=" & iID

				call executeThis(strsql)
				
				Response.Cookies(strCookieURL & "lastviewid") = cLng(Request.QueryString("id"))
				Response.Cookies(strCookieURL & "lastviewid").Expires = dateadd("d",1,now())
			end if
		end if

		'
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY "
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_GOLD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_REP"		
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_GTOTAL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RTOTAL"		
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_ZIP " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_GLOW " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE" 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS" 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1, " & strMemberTablePrefix & "MEMBERS.M_LINK2 "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE, " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX, " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES, " & strMemberTablePrefix & "MEMBERS.M_LNEWS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE, " & strMemberTablePrefix & "MEMBERS.M_BIO "
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PAGE_VIEWS, " & strMemberTablePrefix & "MEMBERS.M_MSN"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID=" & iID

		set rs = my_Conn.Execute(strSql)
		
		if rs.eof then
		  showUpdResult("Member not found!")
		else		
		strMyHobbies = rs("M_HOBBIES")
		strMyLNews = rs("M_LNEWS")
		strMyQuote = rs("M_QUOTE")
		strMyBio = rs("M_BIO")
		intPMstatus = rs("M_PMSTATUS")
		if rs("M_PMRECEIVE") = 0 then
		  intPMstatus = 0
		end if
		
		strColspan = " colspan=""2"" "
		strIMURL1 = "javascript:openWindow('"
		strIMURL5 = "javascript:openWindow5('"
		strIMURL2 = "')"
		
		strMemberDays = DateDiff("d", chkDate2(rs("M_DATE")), chkDate2(strCurDateString))
		'strMemberDays = DateDiff("d", chkDate2(strCurDateString), chkDate2(rs("M_DATE")))
		'response.Write("strMemberDays: " & strMemberDays)
		if strMemberDays = 0 then strMemberDays = 1
		strMemberPostsperDay = round(rs("M_POSTS")/strMemberDays,2)
		set rsposts = my_Conn.Execute("SELECT P_COUNT FROM " & strTablePrefix & "TOTALS")
		if (rsposts("P_COUNT")) <> 0 then
			strMemberPercentPosts = (round(rs("M_POSTS")/(rsposts("P_COUNT")),2)*100)
		else
			strMemberPercentPosts = 0
		end if
		set rsposts = nothing
	  
spThemeBlock1_open(intSkin)
%><table class="tPlain" width="100%">
	  <tr>
	  	 <td align=center>
	  		<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center>
	  		<tr>
<% if mLev = 4 then %>
		<td align=left class="tTitle">&nbsp;<a href="cp_main.asp?cmd=10&mode=Modify&ID=<% =rs("MEMBER_ID") %>&name=<% =ChkString(rs("M_NAME"),"urlpath") %>"><% =ChkString(rs("M_NAME"),"display") %></a></td>
<% else %>				
				<td align=left class="tTitle">&nbsp;<% =ChkString(rs("M_NAME"),"display")  %></b></td>
<% end if%>
				<td align=right class="tTitle"><%= txtMbrSnce %>:&nbsp;<% =ChkDate2(rs("M_DATE")) %>&nbsp;</td>
			</tr>
			</table>
		</td>
	  </tr>
	  <tr>
	    <td align=left valign=top>

				   <table border="0" width="100%" cellspacing="0" cellpadding="3">								    
<%					if strPicture = "1" then %>
					 <tr>
						<td align=center class="tSubTitle" colspan="2"><b><span class="fTitle">&nbsp;<%= txtMyPhoto %>&nbsp;</b></td>
					 </tr>
					 <tr>
					 <% if strUserMemberID <> iID and hasAccess(1) then %>
 					    <td class="tCellAlt0" align=center width="50%">
<%							if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then %>
								<a href="<% =ChkString(rs("M_PHOTO_URL"), "displayimage")%>"><img src="<% =ChkString(rs("M_PHOTO_URL"), "displayimage")%>" alt="<% =ChkString(rs("M_NAME"),"display") %>" height="150" border=0 hspace="2" vspace="2"></a><br><%= txtClkFullSize %>
<%							else %>
								<img src="images/no_photo.gif" title="<%= txtNoPicAvail %>" alt="<%= txtNoPicAvail %>" width="150" height="150" border="0" hspace="2" vspace="2"></a>
<%							end if %>
						</td>
 					    <td class="tCellAlt0" align=center width="50%">
						<% showMemberGroups(iID) %>
						</td>
					 <% Else %>
 					    <td class="tCellAlt0" align=center colspan="2">
<%							if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then %>
								<a href="<% =ChkString(rs("M_PHOTO_URL"), "displayimage")%>"><img src="<% =ChkString(rs("M_PHOTO_URL"), "displayimage")%>" alt="<% =ChkString(rs("M_NAME"),"display") %>" height="150" border=0 hspace="2" vspace="2"></a><br><%= txtClkFullSize %>
<%							else %>
								<img src="images/no_photo.gif" title="<%= txtNoPicAvail %>" alt="<%= txtNoPicAvail %>" width="150" height="150" border="0" hspace="2" vspace="2"></a>
<%							end if %>
						</td>
					<% End If %>
					</tr>
<%					end if ' strPicture %>
						<tr>
							<td valign="top" align=center colspan="2" class="tSubTitle"><span class="fTitle"><%= txtBasics %></span></td>
						</tr>
						<tr>	
							<td class="tCellAlt0" width="40%" align=right nowrap valign=top><b><%= txtUsrNam %>:&nbsp;</b></td><td class="tCellAlt0">
<b><%= displayName(ChkString(rs("M_NAME"),"display"),"") %></b>
							</td>						  
						</tr>
<%				if strAuthType = "nt" then %>
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtUsrNam %>:&nbsp;</b></td>
							<td class="tCellAlt0"><%= ChkString(rs("M_USERNAME"),"display") %></td>
						</tr>
<%				end if 
				if strFullName = "1" and (Trim(rs("M_FIRSTNAME")) <> ""  or  Trim(rs("M_LASTNAME")) <> "" ) then
%>
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtRNam %>:&nbsp;</b></td><td class="tCellAlt0"><% =ChkString(rs("M_FIRSTNAME"), "display") %>&nbsp;<% =ChkString(rs("M_LASTNAME"), "display") %></td>
						</tr>
<%
				end if
				if (strCity = "1" and Trim(rs("M_CITY")) <> "") or (strCountry = "1" and Trim(rs("M_COUNTRY")) <> "") or (strCountry = "1" and Trim(rs("M_STATE")) <> "") then
%>			
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtLocation %>:&nbsp;</b></td><td class="tCellAlt0">
<%	
'## FLag_SQL - Get Flag from DB
		strSql = "SELECT " & strTablePrefix & "COUNTRIES.CO_FLAG"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS INNER JOIN "
		strSql = strSql & strTablePrefix & "COUNTRIES ON "& strMemberTablePrefix & "MEMBERS.M_COUNTRY ="& strTablePrefix & "COUNTRIES.CO_NAME "
		strSql = strSql & "WHERE "& strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & cLng(iID)
		set rsflag = my_Conn.Execute (strSql)

						Response.Write(rs("M_CITY")) 
							if Trim(rs("M_CITY")) <> "" then
								Response.Write("&nbsp;")
							end if
							if Trim(rs("M_STATE")) <> "" then
								Response.Write(ChkString(rs("M_STATE"), "display") & "<br>")
							end if
							Response.Write(ChkString(rs("M_COUNTRY"), "display")&"  ")
							 
						If not IsNull(rs("M_COUNTRY")) And trim(rs("M_COUNTRY")) <> ""  Then 
                        	If Not IsNull(rsflag("CO_FLAG")) And Trim(rsflag("CO_FLAG")) <> "" And rsflag("CO_FLAG") <> " "  Then %> 	
                               <img src="<% =rsflag ("CO_FLAG") %>" align="absmiddle"  border="0px" hspace="0" >
						 <% end If 
                        End If 
'                             rsflag .close 
                             set rsflag  = nothing 
%>
							
							</td>
						</tr>
<%
				end if
				if (strAge = "1" and Trim(rs("M_AGE")) <> "") then
					UBirthdate = rs("M_AGE")
					UAge = GetAge(UBirthdate)
					%>							
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtAge %>:&nbsp;</b></td><td class="tCellAlt0"><% =UAge%></td>
						</tr>
<%
				end if
				if (strMarStatus = "1" and Trim(rs("M_MARSTATUS")) <> "") then
%>			
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtMarStat %>:&nbsp;</b></td><td class="tCellAlt0"><% = ChkString(rs("M_MARSTATUS"), "display") %></td>
						</tr>
<%
				end if
				if (strSex = "1" and Trim(rs("M_SEX")) <> "") then
%>			
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtSex %>:&nbsp;</b></td>
							<td class="tCellAlt0"><% = ChkString(rs("M_SEX"), "display") %></td>
						</tr>
<%
				end if
				if (strOccupation = "1" and Trim(rs("M_OCCUPATION")) <> "") then
%>						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtOcc %>:&nbsp;</b></td>
							<td class="tCellAlt0"><% = ChkString(rs("M_OCCUPATION"), "dislay") %></td>
						</tr>
<%
				end if
			'response.Write("<tr><td colspan=""2"">Xtest here: </td></tr>")
%>			<% If chkApp("forums","USERS") Then %>
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtTotPosts %>:&nbsp;</b></td><td class="tCellAlt0"><% = ChkString(rs("M_POSTS"), "display")%> (<%= strMemberPercentPosts%>%&nbsp;<%= txtOfTotPosts %> / <% =strMemberPostsperDay %>&nbsp;<%= txtPostsPerDay %>)</td>
						</tr>
			<% End If %>
						<% If showGold = 1 Then %>
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtGold %>:&nbsp;</b></td>
							<td class="tCellAlt0"><% = ChkString(rs("M_GOLD"), "display") %></td>
						</tr>
						<% End If %>
						<% If showRep = 1 Then %>
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtRepPts %>:&nbsp;</b></td><td class="tCellAlt0"><% = ChkString(rs("M_REP"), "display") %></td>
						</tr><% End If %>
						<% If showGames = 1 Then %>						
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtTotGms %>:&nbsp;</b></td><td class="tCellAlt0"><% = ChkString(rs("M_GTOTAL"), "display") %></td>
						</tr><% End If %>
						<tr>
							<td class="tCellAlt0" align=right nowrap valign=top><b><%= txtRfls %>:&nbsp;</b></td><td class="tCellAlt0"><% = ChkString(rs("M_RTOTAL"), "display") %></td>
						</tr>
						<tr>
							<td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtProfViews %>:&nbsp;</b></td><td class="tCellAlt0"><% = rs("M_PAGE_VIEWS") %></td>
						</tr>							
					 <tr>
						<td align="center" class="tSubTitle" colspan="2"><span class="fTitle"><b>&nbsp;<%= txtCtInfo %>&nbsp;</b></span></td></tr>
					 <tr>
           				 <td class="tCellAlt0" align="right" nowrap><b><%= txtEmlAdd %>:&nbsp;</b></td><%				if Trim(rs("M_EMAIL")) <> "" then %>
           				<td class="tCellAlt0" nowrap>&nbsp;<% if rs("M_HIDE_EMAIL") = 1 and not mlev=4 then %><%= txtHidByMbr %><% else %><a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><%= txtClkSend %>&nbsp;<% =ChkString(rs("M_NAME"), "display") %>&nbsp;<%= txtAnEml %></a><% end if %>&nbsp;</td>
<%				else %>
           				<td class="tCellAlt0"><%= txtNoEmlSpec %></td>
<%				end if %>
         				</tr>
<%				if strMSN = "1" then
					if Trim(rs("M_MSN")) <> "" then %>
					<tr>
						<td class="tCellAlt0" align=right nowrap><b><%= txtMSN %>:&nbsp;</b></td>
						<td class="tCellAlt0"><a href="<% =strIMURL1 %>pop_portal.asp?cmd=7&mode=3&msn=<% =ChkString(replace(rs("M_MSN"),"@","[no-spam]@"), "displayimage") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %><% =strIMURL2 %>"><img src="images/icons/icon_msn.gif" border="0" align="absmiddle">&nbsp;<% =ChkString(replace(rs("M_MSN"),"@","[no-spam]@"), "display") %></a>&nbsp;</td>
					</tr>
<%					end if
				end if %>
<%				if strICQ = "1" then
					if Trim(rs("M_ICQ")) <> "" then %>
					<tr>
						<td class="tCellAlt0" align=right nowrap><b><%= txtICQ %>:&nbsp;</b></td>
						<td class="tCellAlt0"><a href="<% =strIMURL1 %>pop_portal.asp?cmd=7&mode=1&ICQ=<% =ChkString(rs("M_ICQ"), "urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %><% =strIMURL2 %>"><img src="http://web.icq.com/whitepages/online?icq=<% = ChkString(rs("M_ICQ"),"display")  %>&img=5" border="0" align="absmiddle"><% =ChkString(rs("M_ICQ"), "display") %></a>&nbsp;</td>
					</tr>
<%					end if
				end if
				if strYAHOO = "1" then
					if Trim(rs("M_YAHOO")) <> "" then %>
					<tr>
						<td class="tCellAlt0" align=right nowrap><b><%= txtYhoIM %>:&nbsp;</b></td>
						<td class="tCellAlt0"><a href="<% =strIMURL5 %>http://edit.yahoo.com/config/send_webmesg?.target=<% =ChkString(rs("M_YAHOO"), "urlpath") %>&.src=pg<% =strIMURL2 %>"><img border=0 src="http://opi.yahoo.com/online?u=<% =ChkString(rs("M_YAHOO"), "urlpath") %>&m=g&t=2"></a>&nbsp;</td>
					</tr>
<%					end if
				end if
				if strAIM = "1" then
					if Trim(rs("M_AIM")) <> "" then %>
					<tr>
						<td class="tCellAlt0" align=right nowrap><b><%= txtAIM %>:&nbsp;</b></td>
						<td class="tCellAlt0"><a href="<% =strIMURL1 %>pop_portal.asp?cmd=7&mode=2&AIM=<% =ChkString(rs("M_AIM"), "urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %><% =strIMURL2 %>"><% =ChkString(rs("M_AIM"), "urlpath") %></a>&nbsp;</td>
					</tr>
<%					end if
				end if  
%>
<% 
if chkApp("PM","USERS") and intPMstatus = 1 then %>
<tr>
<td class="tCellAlt0" align=right nowrap><b><%= txtPvtMessg %>:</b></td><td class="tCellAlt0">
&nbsp;<a href="Javascript:;" onclick="Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=<%= getmemberid(rs("M_NAME")) %>');"><%= replace(txtSndPvtMsg,"[%member%]",ChkString(rs("M_NAME"), "display")) %></a></td>
</tr><% 
end if
%>						
<%	
				if (strHomepage + strFavLinks) > 0 then  %>
				<tr>
					<td align=center class="tSubTitle" colspan="2">
					<b><span class="fTitle"><%= txtLinks %>&nbsp;</b></td>			
<%					if strHomepage = "1" then %>
						<tr>
							<td class="tCellAlt0" align=right nowrap width="10%"><b><%= txtHomePg %>:&nbsp;</b></td>
<%							if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then %>
							<td class="tCellAlt0"><a href="<% =rs("M_Homepage") %>" target="_blank"><% =rs("M_Homepage") %></a>&nbsp;</td>
<%							else %>
							<td class="tCellAlt0"><%= txtNoHmPg %>...</td>
<%							end if %>
						</tr>
<%					end if
						
					if strFavLinks = "1" then %>
						<tr>
							<td class="tCellAlt0" align=right valign="top" nowrap><b><%= txtClLnks %>:&nbsp;</b></td>
<%						if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then %>
							<td class="tCellAlt0">
							<a href="<% =rs("M_LINK1") %>" target="_Blank"><% =rs("M_LINK1") %></a>&nbsp;
<%						  if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then %>
							<br><a href="<%=rs("M_LINK2")%>" target="_Blank"><%=rs("M_LINK2")%></a>&nbsp;
<%						  end if %>
							</td></tr>
<%						else %>
							<td class="tCellAlt0"><%= txtNoLnksSp %>...</td>
						
<%						end if %>
						</tr>		
<%					end if 

				end if ' links
				if (strBio + strHobbies + strLNews + strQuote) > 0 then %>			
				<tr>
					<td align=center class="tSubTitle" colspan="2"><%= txtMAbtMe %></td></tr>
<%				if strBio = "1" then  %>
				<tr>				
					<td class="tCellAlt0" valign=top align=right nowrap width="10%">
					<b><% =strVar1%>:&nbsp;</b>
					</td>	
					<td class="tCellAlt0" valign=top>
					<% if IsNull(strMybio) or trim(strMyBio) = "" then Response.Write("-") else Response.Write(formatStr(strMyBio)) %>
					</td>
				</tr>
<%				end if
				if strHobbies = "1" then  %>
				<tr>
					<td class="tCellAlt0" valign=top align=right nowrap width="10%">
					<b><% =strVar2%>:&nbsp;</b>
					</td>
					<td class="tCellAlt0" >
					<% if IsNull(strMyHobbies)  or trim(strMyHobbies) = "" then Response.Write("-") else Response.Write(formatStr(strMyHobbies)) %>
					</td>
				</tr>
<%				end if
				if strLNews = "1" then  %>
				<tr>
					<td class="tCellAlt0" valign=top align=right nowrap width="10%">
					<b><% =strVar3 %>:&nbsp;</b>
					</td>
					<td class="tCellAlt0" valign=top>
					<% if IsNull(strMyLNews) or trim(strMyLNews) = "" then Response.Write("-") else Response.Write(formatStr(strMyLNews)) %>
					</td>
				</tr>
<%				end if
				if strQuote = "1" then  %>
				<tr>
					<td align=right class="tCellAlt0" nowrap width="10%" valign=top>
					<b><% =strVar4 %>:&nbsp;</b></td>
					<td class="tCellAlt0" valign=top>
					<% if IsNull(strMyQuote) or Trim(strMyQuote) = "" then Response.Write("-") else Response.Write(formatStr(strMyQuote)) %></td>
				</tr>
<%				end if
			end if
			set rs = nothing
%>		
			</table>			
</td>
</tr></table>
<%
spThemeBlock1_close(intSkin)
end if	%>

<table border="0" width="100%" cellspacing="0" cellpadding="3" valign=top>
	<tr>
		<td align=center nowrap>
			<br>
			<a href="JavaScript: onClick= history.go(-1) "><%= txtBack %></a></p>
			<p>&nbsp;</p></td></tr>
	</table>			
<!-- </td>
</tr></table> -->
<%
end sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub editProfile()
  spThemeTitle = txtEditProf
  spThemeBlock1_open(intSkin)
  select case request.QueryString("mode")
    case "EditIt"
	  'response.Write("Mode: EditIt") 
	  Err_Msg = ""
    
      if (trim(Request.Form("B_Month")) <> "" and trim(Request.Form("B_Day")) <> "" and trim(Request.Form("B_Year")) <> "" )  then
         formbirthdate = chkString(Request.Form("B_Month"),"sqlstring") & "/" & chkString(Request.Form("B_Day"),"sqlstring") & "/" & chkString(Request.Form("B_Year"),"sqlstring")
         ' Check to see if birthdate is a valid date
	    If NOT IsDate(formbirthdate) Then
		  Err_Msg = Err_Msg & "<li>" & txtValBday & "!</li>" 
	    End If 
	    If IsDate(formbirthdate) then
          if CDate(formbirthdate) > CDate(strForumTimeAdjust) then
          Err_Msg = Err_Msg & "<li>" & txtBdayPrior & "</li>"
          end if
        end if
	  else
	  	formbirthdate = " "	
      end if
		
	  if Request.Form("Name") = "" then 
		Err_Msg = Err_Msg & "<li>" & txtChoseUsrNam & "</li>"
	  end if
	  if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) or (Instr(Request.Form("Name"), ",") > 0) or (Instr(Request.Form("Name"), "&") > 0) or (Instr(Request.Form("Name"), "#") > 0) or (Instr(Request.Form("Name"), "'") > 0) then
			Err_Msg = Err_Msg & "<li>" & txtCharsNotAllow & "<br> &gt; " & txtand & " &lt;</li>"
	  end if
	  if strAuthType = "db" then
	     if Len(Request.Form("Password")) > 0 then
			if Len(Request.Form("Password")) > 25 or Len(Request.Form("Password")) < 5 then 
				Err_Msg = Err_Msg & "<li>" & txtUPassLen & "</li>" 
			end if
			if trim(Request.Form("Password")) <> "" then
			  if trim(Request.Form("Password")) <> trim(Request.Form("Password2")) then 
				Err_Msg = Err_Msg & "<li>" & txtPassNoMatch & "</li>"
			  end if
			end if
			if (Instr(Request.Form("Password"), ">") > 0 ) or (Instr(Request.Form("Password"), "<") > 0) or (Instr(Request.Form("Password"), ",") > 0) or (Instr(Request.Form("Password"), "&") > 0) or (Instr(Request.Form("Password"), "#") > 0) or (Instr(Request.Form("Password"), "'") > 0) then
				Err_Msg = Err_Msg & "<li>" & txtCharsNotAllow & "</li>"
			end if
		 end if
	  end if
	  if Request.Form("Email") = "" then 
		Err_Msg = Err_Msg & "<li>" & txtErNoEmlAdd & "</li>"
	  end if
	  if EmailField(Request.Form("Email")) = 0 then 
		Err_Msg = Err_Msg & "<li>" & txtErValEml & "</li>"
	  end if
	  if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
		Err_Msg = Err_Msg & "<li>" & txtPrefixUrl & "</li>"
	  end if
	  sNewEmail = false
	  if strUniqueEmail = "1" then
		if ((lcase(Request.Form("Email")) = lcase(Request.Form("Email2"))) and (lcase(Request.Form("Email")) <> lcase(Request.Form("Email3")))) then
			strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE M_EMAIL='"& Trim(chkString(Request.Form("Email"),"sqlstring")) &"'"
			set rs = my_Conn.Execute (strSql)
			if rs.BOF and rs.EOF then 
				' Do Nothing
			Else 
				Err_Msg = Err_Msg & "<li>" & txtEmlInUse & "</li>"
			end if
			rs.close
			set rs = nothing
			
			if lcase(strEmail) = "1" and Err_Msg = "" and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8) then
				verKey = GetKey("sendemail")
				sNewEmail = true
			end if
		end if
	  else
		if ((lcase(Request.Form("Email")) = lcase(Request.Form("Email2"))) and (lcase(Request.Form("Email")) <> lcase(Request.Form("Email3")))) and lcase(strEmail) = "1" and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8) then
			verKey = GetKey("sendemail")
			sNewEmail = true
		end if
	  end if
	  
	  if Err_Msg = "" then
		
		if Trim(Request.Form("Homepage")) <> "" and lcase(trim(Request.Form("Homepage"))) <> "http://" and Trim(lcase(Request.Form("Homepage"))) <> "https://" then
			regHomepage = ChkString(Request.Form("Homepage"),"display")
		else
			regHomepage = " "
		end if
		if Trim(Request.Form("LINK1")) <> "" and lcase(trim(Request.Form("LINK1"))) <> "http://" and Trim(lcase(Request.Form("LINK1"))) <> "https://" then
			regLink1 = ChkString(Request.Form("LINK1"),"display")
		else
			regLink1 = " "
		end if
		if Trim(Request.Form("LINK2")) <> "" and lcase(trim(Request.Form("LINK2"))) <> "http://" and Trim(lcase(Request.Form("LINK2"))) <> "https://" then
			regLink2 = ChkString(Request.Form("LINK2"),"display")
		else
			regLink2 = " "
		end if
		if Trim(Request.Form("PHOTO_URL")) <> "" and lcase(trim(Request.Form("PHOTO_URL"))) <> "http://" and Trim(lcase(Request.Form("PHOTO_URL"))) <> "https://" then
			regPhoto_URL = ChkString(Request.Form("Photo_URL"),"display")
		else
			regPhoto_URL = " "
		end if

		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		if trim(Request.Form("Password")) = "" then
			strSql = strSql & " SET M_COUNTRY  = '" & ChkString(Request.Form("Country"),"sqlstring")  & " ', "
		else
			strSql = strSql & " SET M_PASSWORD = '" & pEncrypt(pEnPrefix & ChkString(Request.Form("Password"),"sqlstring")) & "', "
			strSql = strSql & "     M_COUNTRY  = '" & ChkString(Request.Form("Country"),"sqlstring")  & " ', "
		end if
		strSql = strSql & "     M_RECMAIL  = '" & ChkString(Request.Form("RECMAIL"),"sqlstring")  & "', "
			
		if strICQ = "1" then
			strSql = strSql & "     M_ICQ = '" & ChkString(Request.Form("ICQ"),"sqlstring") & " ', "
		end if
		if strYAHOO = "1" then
			strSql = strSql & "     M_YAHOO = '" & ChkString(Request.Form("YAHOO"),"sqlstring") & " ', "
		end if
		if strAIM = "1" then
			strSql = strSql & "     M_AIM = '" & ChkString(Request.Form("AIM"),"sqlstring") & " ', "
		end if
		if strHOMEPAGE = "1" then
			sTmpStr = ChkString(Trim(regHomepage),"sqlstring")
			strSql = strSql & "     M_Homepage = '" & sTmpStr & " ', "
		end if
		strSql = strSql & "     M_SIG = '" & ChkString(Request.Form("Sig"),"message") & " ', "
		if (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 8) then
			strSql = strSql & "     M_NEWEMAIL = '" & ChkString(Request.Form("Email"),"SQLString") & "' "
		else
			strSql = strSql & "     M_EMAIL = '" & ChkString(Request.Form("Email"),"SQLString") & "' "
		end if
		strSql = strSql & ", 	M_KEY = '" & chkString(verKey,"SQLString") & "'"
		if strfullName = "1" then
			strSql = strSql & ",	M_FIRSTNAME = '" & ChkString(Request.Form("FirstName"),"sqlstring") & "'" 
			strSql = strSql & ",	M_LASTNAME  = '" & ChkString(Request.Form("LastName"),"sqlstring") & "'"  
		end if
		if strCity = "1" then
			strsql = strsql & ",	M_CITY = '" & ChkString(Request.Form("City"),"sqlstring") & "'"  
		end if
		if strState = "1" then
			strsql = strsql & ",	M_STATE = '" & ChkString(Request.Form("State"),"sqlstring") & "'" 
		end if
		if strZip = "1" then
		  if trim(Request.Form("Zipcode")) <> "" then
			strsql = strsql & ",	M_ZIP = '" & ChkString(Request.Form("Zipcode"),"sqlstring") & "'" 
		  else
			strsql = strsql & ",	M_ZIP = ''" 
		  end if
		end if
		strsql = strsql & ", M_HIDE_EMAIL = '" & ChkString(Request.Form("HideMail"),"sqlstring") & "'"  
		if strPicture = "1" then
		  strsql = strsql & ", M_PHOTO_URL = '" & ChkString(Request.Form("Photo_URL"),"display") & "'"  
		end if
		  strsql = strsql & ", M_AVATAR_URL = '" & ChkString(Request.Form("Avatar_URL"),"display") & "'"  			
		if strFavLinks = "1" then
		  strsql = strsql & ", M_LINK1 = '" & ChkString(Request.Form("LINK1"),"display") & "'" 
		  strSql = strSql & ", M_LINK2 = '" & ChkString(Request.Form("LINK2"),"display") & "'" 
		end if
		if strAge = "1" then
		  strSql = strsql & ", M_AGE = '" & formbirthdate & "'"
		end if
		if strMarStatus = "1" then
		  strSql = strSql & ", M_MARSTATUS = '" & ChkString(Request.Form("MarStatus"),"sqlstring") & "'" 
		end if
		if strSex = "1" then
		  strSql = strsql & ", M_SEX = '" & ChkString(Request.Form("Sex"),"sqlstring") & "'" 
		end if
		if strOccupation = "1" then
		  strSql = strSql & ", M_OCCUPATION = '" & ChkString(Request.Form("Occupation"),"sqlstring") & "'" 
		end if
		if strBio = "1" then
		  strSql = strSql & ", M_BIO = '" & ChkString(Request.Form("Bio"),"sqlstring") & "'" 
		end if
		if strHobbies = "1" then
		  strSql = strSql & ", M_HOBBIES = '" & ChkString(Request.Form("Hobbies"),"sqlstring") & "'" 
		end if
		if strLNews = "1" then
		  strsql = strsql & ", M_LNEWS = '" & ChkString(Request.Form("LNews"),"sqlstring") & "'" 
		end if
		if strQuote = "1" then
		  strSql = strSql & ",	M_QUOTE = '" & ChkString(Request.Form("Quote"),"sqlstring") & "'" 
		end if
		if strMSN = "1" then
		  strSql = strSql & ",	M_MSN = '" & ChkString(Request.Form("MSN"),"sqlstring") & "'" 
		end if
		strSql = strSql & " WHERE M_NAME = '" & chkstring(request.form("Name"), "sqlstring") & "' "
		if strAuthType = "db" then 
		  strSql = strSql & " AND M_PASSWORD = '"& chkstring(request.form("Password-d"), "sqlstring") &"'"
		end IF
		'response.Write(strSql & "<br>")			
		executeThis(strsql)
			
		regHomepage = ""
			
		tmpResult = "<span class=""fTitle"">" & txtProfUpd & "</span></p>"
		%>
		
		<%
		if sNewEmail = true and lcase(strEmail) = "1"  and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 8) then
			tmpResult = tmpResult & "<p align=""center"">" & txtEmlHasChgd & "<br><br></p>"
		else 
			tmpResult = tmpResult & "<meta http-equiv=""Refresh"" content=""2; URL=cp_main.asp"">"
		    tmpResult = tmpResult & "<p align=""center""><a href=""cp_main.asp"">" & txtBack & "</a></p>"
		end if
		response.write(tmpResult & "<br>")
	  else 'error message
	    %><p>&nbsp;</p>
		<p align=center><span class="fTitle"><%= txtThereIsProb %></span></p>
		<table align="center"><tr><td align="center">
		<ul><% =Err_Msg %></ul>
	    </td></tr></table>
		<p align=center><a href="JavaScript:history.go(-1)"><%= txtBack %></a></p>
		<p>&nbsp;</p>
		<%
	  end if
	  
    case "goEdit"
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME" 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_EMAIL" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LASTNAME" 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_ZIP" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AVATAR_URL" 		 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK2"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION" 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LNEWS" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_BIO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RECMAIL" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_MSN"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
		strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"
		if strAuthType = "db" then
			strSql = strSql & " AND M_PASSWORD = '" & pEncrypt(pEnPrefix & Request.Form("Password")) & "'"
		end if
		QuoteOk = ChkQuoteOk(STRdbntUserName)
		set rs = my_Conn.Execute(strSql)
'Response.write strSQL 
'response.end
		if rs.BOF and rs.EOF and QuoteOk then 
		  %>
		  <p align="center"><span class="fTitle"><%= txtBadLogin1 %></span></p>
		  <p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoAuth %></a></p>
		  <%
		else
		  If (SecImage > 1 and DoSecImage(Ucase(request.form("SecCode"))) <> 1) Then	%>
  		  	<p align=center><span class="fTitle"><%= txtBadSecCode %></span></p>
			<p align="center"><a href="default.asp"><%= txtGoAuth %></a></p>
		  <%	
		  Else	
			'## Display Edit Profile Page
		  %>
			<p align=center><span class="fTitle"><%= txtEditMProf %></span></p>
			<p align=center>
			<form action="cp_main.asp?cmd=9&mode=EditIt" method="Post" id="Form1" name="Form1">
			<input name="Refer" type="hidden" value="<% =chkString(Request.Form("Refer"),"refer") %>">
			<!-- #include file="inc_profile.asp" -->
			</form>
			</p><br>
		  <%
		  end if 
		end if
		
	case else %>
	  <script language="JavaScript" type="text/JavaScript">
	  function focuspass() { document.forms.Form1.Password.focus(); }
	  window.onload=focuspass;
	  </script>
	  <form action="cp_main.asp?cmd=9&mode=goEdit&id=<% =cLng(Request.QueryString("id"))%>" method="post" name="Form1" id="Form1">
	  <p align="center"><span class="fTitle"><%= txtEditProf %></span></p><p align="center">
	  <input name="Refer" type="hidden" value="<% =chkString(Request.ServerVariables("HTTP_REFERER"), "refer") %>">
	  <%= txtProfUpToDat %><br>
<%		if strAuthType = "nt" then %>
		  <%= txtNTlogin %><br><br>
<%		else %>
		  <%= txtPlzLogAgin %><br><br>	
<%		end if %>	
		</p>
	  <center><div style="width:300px;">
	  <table class="tPlain">
<% 	  if strAuthType = "nt" then %>
		<TR>
			<TD align=right nowrap><b><%= txtUAcct %>:</b></td>
			<TD><span class="fTitle"><% =Session(strCookieURL & "userid") %></span></b></td>
		</TR>
<%	  else %>	
	 	<TR>
	 	    <TD align=right nowrap><b><%= txtUsrNam %>:&nbsp;</b></td>
	 	    <TD><input name="Name" size="15" value="<% =chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>"></td>
	 	</TR>
	 	<TR>
	 	    <TD align=right nowrap><b><%= txtPass %>:&nbsp;</b></td>
	 	    <TD><input name="Password" type="Password" size="15">
	 	    </td>
	 	</TR>
<% 		If SecImage >1 Then %>
  		<TR>         
	 		<TD  class="tCellAlt1" align=center colspan="2" > 		
			<img src="includes/securelog/image.asp" />
	 		</td>	 	
		</TR> 
  		<TR>
	 	    <TD  class="tCellAlt1" align=right nowrap><b><%= txtSecCode %>:&nbsp;</b></td> 	 	    
	 	    <TD  class="tCellAlt1"><input  name="secCode" size="15" value="" onFocus="javascript:this.value='';"></td>
  		</TR>	 	  
<%		End If %> 
<%	  end If 'db type check %>
	 	<TR>
	 		<TD align=center colspan=2><input type="submit" value="<%= txtSubmit %>" class="button"></td></TR>
	</table></div></center>
	</form><br>&nbsp;
	<%
  end select
  spThemeBlock1_close(intSkin)
end sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


sub editMemberProfile()
select case Request.QueryString("mode")
	case "Modify"
	  %>
	  <form action="cp_main.asp?cmd=10&mode=goModify" method="post" id="Form1" name="Form1">
	  <input type="hidden" name="MEMBER_ID" value="<% =cLng(Request.QueryString("id")) %>">
	  <%
	  strBlockMessage = "<p align=center><span class=""fTitle"">" & txtEditMProf & "</span></p>"
	  strBlockMessage = strBlockMessage & "<p align=center><span class=""fAlert""><b>" & ucase(txtNote) & ":</b></span>"
	  strBlockMessage = strBlockMessage & txtAdminCanMod & "</p><p>&nbsp;</p>"
	
	  showPasswordBlock 1,txtEditMProf,strBlockMessage,0,0
	  %>
	  </form>
	  <%
	case "goModify"
		mName = chkString(Request.Form("User"),"display")
		mPassword = pEncrypt(pEnPrefix & request.Form("Pass"))
		m_id = clng(request.Form("MEMBER_ID"))
	    mMlev = getMlev(mName, mPassword)
		
		strMsg = txtNoPermViewPg
		doit = ""
		if m_id <> strUserMemberID and mMlev = 4 and mLev = 4 then
		  thisUserS = chkIsSuperAdmin(1,strUserMemberID)
		  targUserS = chkIsSuperAdmin(1,m_id)
		  targUserA = chkIsAdmin(m_id)
		  if targUserS = 1 and thisUserS = 1 then
		    doit = ""
			strMsg = txtNoEditSA
			'response.Write(strMsg & "<br>")
		  elseif targUserS > thisUserS then
		    doit = ""
			strMsg = txtNoEditSA
			'response.Write(strMsg & "<br>")
		  elseif targUserA = 1 and thisUserS = 0 then
		    doit = ""
			strMsg = txtNoEditOAd
			'response.Write(strMsg & "<br>")
		  else
			doit = "ok"
			strMsg = txtEditMProf
			'response.Write(strMsg & "<br>")
		  end if
		elseif mMlev = 4 and mLev = 4 then
			doit = "ok"
			'strMsg = "You are editing yourself or another member"
			'response.Write(strMsg & "<br>")
		end if
						
		if doit = "ok" then  '## is Member
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LASTNAME " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_GOLD"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_REP"						
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RNAME"	
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RTOTAL"							
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_ZIP "
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AVATAR_URL" 			
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE" 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK2 "
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LNEWS " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_BIO "		
	   		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RECMAIL" 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MSN"
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE MEMBER_ID = " & chkstring(request.form("MEMBER_ID"), "sqlstring")

		  set rs = my_Conn.Execute(strSql)
		  if rs.eof then
	  	    tmpResult = ""
	  	    tmpResult = tmpResult & "<p align=center><span class=""fAlert""><b>" & txtERROR & "</b></span></p>"
	  	    tmpResult = tmpResult & "<p><span class=""fTitle"">" & txtMemNoFnd & "</span>"
	  	    showUpdResult(tmpResult)
		  else
			'## Display Edit Profile Page
			spThemeTitle = txtEditMProf
			spThemeBlock1_open(intSkin) %>
			<p align=center><span class="fTitle"><%= txtEditMProf %></span></p>
			<p align=center><form action="cp_main.asp?cmd=10&mode=ModifyIt&id=<% =chkString(Request.Form("MEMBER_ID"),"sqlstring")%>" method="Post" id="Form1" name="Form1">
			<input type=hidden name="User" value="<% =chkString(Request.Form("User"),"sqlstring") %>">
			<input type=hidden name="Pass" value="<% =chkString(Request.Form("Pass"),"sqlstring") %>">
			<input name="id" type="hidden" value="<%= cLng(Request.Form("MEMBER_ID")) %>">
			<!-- #include file="inc_profile.asp" -->
			</form></p>
			<%
			spThemeBlock1_close(intSkin)
		  end if
		  set rs = nothing			
		else 
	  	  tmpResult = ""
	  	  tmpResult = tmpResult & "<p align=center><span class=""fAlert""><b>" & txtERROR & "</b></span></p>"
	  	  tmpResult = tmpResult & "<p><span class=""fTitle"">" & strMsg & "</span>"
	  	  showUpdResult(tmpResult)%>
<%
		end if 
		
	case "ModifyIt"
	  Err_Msg = ""
	  if Request.Form("Name") = "" then 
		Err_Msg = Err_Msg & "<li>" & txtChoseUsrNam & "</li>"
	  end if
	  if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) then
		Err_Msg = Err_Msg & "<li>" & txtCharsNotAllow & " = &gt; " & txtand & " &lt; </li>"
	  end if
				'
	  strSql = "SELECT M_NAME FROM " & strMemberTablePrefix & "MEMBERS "
	  strSql = strSql & " WHERE M_NAME = '" & Trim(chkString(Request.Form("Name"),"sqlstring")) &"' "
	  strSql = strSql & " AND MEMBER_ID <> " & Trim(chkString(Request.Form("Member_ID"),"sqlstring")) &" "
	  set rs = my_Conn.Execute(strSql)	

	  if rs.BOF and rs.EOF then 
		' Do Nothing
	  else 
		Err_Msg = Err_Msg & "<li>" & txtChsAnother & "</li>"
	  end if
						
	  rs.close
	  set rs = nothing
	  if strAuthType = "db" then
		if (Len(Request.Form("Password")) > 25 or Len(Request.Form("Password")) < 5) and Len(Request.Form("Password")) > 0 then 
		  Err_Msg = Err_Msg & "<li>" & txtUPassLen & "</li>" 
		end if
	  end if		   
    
      if (trim(Request.Form("B_Month")) <> "" and trim(Request.Form("B_Day")) <> "" and trim(Request.Form("B_Year")) <> "" )  then
         formbirthdate = ChkString(Request.Form("B_Month"),"sqlstring") & "/" & ChkString(Request.Form("B_Day"),"sqlstring") & "/" & ChkString(Request.Form("B_Year"),"sqlstring")
         ' Check to see if birthdate is a valid date
	    If NOT IsDate(formbirthdate) Then
		  Err_Msg = Err_Msg & "<li>" & txtValBday & "</li>" 
	    End If 
	    If IsDate(formbirthdate) then
          if CDate(formbirthdate) > CDate(strForumTimeAdjust) then
          Err_Msg = Err_Msg & "<li>" & txtBdayPrior & "</li>"
          end if
        end if
	  else
	  	formbirthdate = " "	
      end if
	  
	  if Request.Form("Email") = "" then 
		Err_Msg = Err_Msg & "<li>" & txtErNoEmlAdd & "</li>"
	  end if
	  if EmailField(Request.Form("Email")) = 0 then 
		Err_Msg = Err_Msg & "<li>" & txtErValEml & "</li>"
	  end if
	  if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
		Err_Msg = Err_Msg & "<li>" & txtPrefixUrl & "</li>"
	  end if
	  sNewEmail = false
	  if strUniqueEmail = "1" then
		if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) then
		  strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
		  strSql = strSql & " WHERE M_EMAIL = '" & Trim(chkString(Request.Form("Email"),"sqlstring")) &"'"
		  set rs = my_Conn.Execute (strSql)
		  if rs.BOF and rs.EOF then
			'## Do Nothing - proceed
		  Else
			Err_Msg = Err_Msg & "<li>" & txtEmlInUse & "</li>"
		  end if
		  rs.close
		  set rs = nothing
		  if lcase(strEmail) = "1" and Err_Msg = ""  and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 8) then
			verKey = GetKey("sendemail")
			sNewEmail = true
		  end if
		end if
	  else
		if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) and lcase(strEmail) = "1"  and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 8) then
		  verKey = GetKey("sendemail")
		  sNewEmail = true
		end if
	  end if
	  if Len(Request.Form("Sig")) > 500 then
		'Err_Msg = Err_Msg & "<li>" & txtSigTooLng & "<br>"
		'Err_Msg = Err_Msg & txtLenIs & "<b>" & Len(Request.Form("Sig")) & "</b>.</li>"
	  end if
	  
	  if Err_Msg = "" then '## it is ok to update the profile
		if Trim(Request.Form("Homepage")) <> "" and lcase(trim(Request.Form("Homepage"))) <> "http://" and Trim(lcase(Request.Form("Homepage"))) <> "https://" then
		  regHomepage = ChkString(Request.Form("Homepage"),"display")
		else
		  regHomepage = " "
		end if
		if Trim(Request.Form("LINK1")) <> "" and lcase(trim(Request.Form("LINK1"))) <> "http://" and Trim(lcase(Request.Form("LINK1"))) <> "https://" then
			regLink1 = ChkString(Request.Form("LINK1"),"display")
		else
			regLink1 = " "
		end if
		if Trim(Request.Form("LINK2")) <> "" and lcase(trim(Request.Form("LINK2"))) <> "http://" and Trim(lcase(Request.Form("LINK2"))) <> "https://" then
			regLink2 = ChkString(Request.Form("LINK2"),"display")
		else
			regLink2 = " "
		end if
		if Trim(Request.Form("PHOTO_URL")) <> "" and lcase(trim(Request.Form("PHOTO_URL"))) <> "http://" and Trim(lcase(Request.Form("PHOTO_URL"))) <> "https://" then
			regPhoto_URL = ChkString(Request.Form("Photo_URL"),"display")
		else
			regPhoto_URL = " "
		end if
			
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_NAME = '" & ChkString(Request.Form("Name"),"sqlstring") & "'"
					
		if strAuthType = "nt" then
			strSql = strSql & ", M_USERNAME = '" & ChkString(Request.Form("Account"),"sqlstring") & "'"
		else
		  if strAuthType = "db" and trim(Request.Form("Password")) <> "" then
			strSql = strSql & ", M_PASSWORD = '" & pEncrypt(pEnPrefix & ChkString(Request.Form("Password"),"sqlstring")) & "'"
		  end if
		end if
		if (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 8) then
		  strSql = strSql & ", M_NEWEMAIL = '" & chkString(Request.Form("Email"),"SQLString") & "'"
		else
		  strSql = strSql & ", M_EMAIL = '" & chkString(Request.Form("Email"),"SQLString") & "'"
		end if
		strSql = strSql & ", M_KEY = '" & chkString(verKey,"SQLString") & "'"
		strSql = strSql & ", M_RECMAIL  = '" & ChkString(Request.Form("RECMAIL"),"sqlstring")  & "'"
		strSql = strSql & ", M_TITLE = '" & ChkString(Request.Form("Title"),"sqlstring") & " '"
		strSql = strSql & ", M_POSTS = " & ChkString(Request.Form("Posts"),"sqlstring") & " "
		strSql = strSql & ", M_GOLD = " & ChkString(Request.Form("Gold"),"sqlstring") & " "
		strSql = strSql & ", M_REP = " & ChkString(Request.Form("Rep"),"sqlstring") & " "										
		strSql = strSql & ", M_RTOTAL = " & ChkString(Request.Form("Referrals"),"sqlstring") & " "										
		strSql = strSql & ", M_RNAME = '" & ChkString(Request.Form("Referrer"),"sqlstring") & "'"																				
		strSql = strSql & ", M_COUNTRY = '" & ChkString(Request.Form("Country"),"sqlstring") & " '"
					
		if strICQ = "1" then
			strSql = strSql & ", M_ICQ = '" & ChkString(Request.Form("ICQ"),"sqlstring") & " '"
		end if
		if strYAHOO = "1" then
			strSql = strSql & ", M_YAHOO = '" & ChkString(Request.Form("YAHOO"),"sqlstring") & " '"
		end if
		if strAIM = "1" then
			strSql = strSql & ", M_AIM = '" & ChkString(Request.Form("AIM"),"sqlstring") & " '"
		end if
		if strHOMEPAGE = "1" then
			strSql = strSql & ", M_HOMEPAGE = '" & ChkString(Request.Form("Homepage"),"sqlstring" ) & " '"
		end if
		strSql = strSql & ", M_SIG = '" & ChkString(Request.Form("Sig"),"message") & " '"
		strSql = strSql & ", M_LEVEL = " & ChkString(Request.Form("Level"),"")
		if strfullName = "1" then
			strSql = strSql & ", M_FIRSTNAME = '" & ChkString(Request.Form("FirstName"),"sqlstring") & "'" 
			strSql = strSql & ", M_LASTNAME  = '" & ChkString(Request.Form("LastName"),"sqlstring") & "'"  
		end if
		if strCity = "1" then
			strsql = strsql & ", M_CITY = '" & ChkString(Request.Form("City"),"sqlstring") & "'"  
		end if
		if strState = "1" then
			strsql = strsql & ", M_STATE = '" & ChkString(Request.Form("State"),"sqlstring") & "'" 
		end if
		if strZip = "1" then
			strsql = strsql & ", M_ZIP = '" & ChkString(Request.Form("Zipcode"),"sqlstring") & "'" 
		end if
		'strsql = strsql & ",	M_HIDE_EMAIL = '" & ChkString(Request.Form("HideMail"),"") & "'"  
		if strPicture = "1" then
			strsql = strsql & ", M_PHOTO_URL = '" & ChkString(Request.Form("Photo_URL"),"display") & "'"  
		end if
		strsql = strsql & ", M_AVATAR_URL = '" & ChkString(Request.Form("Avatar_URL"),"display") & "'"  					
		if strFavLinks = "1" then
			strsql = strsql & ", M_LINK1 = '" & ChkString(Request.Form("LINK1"),"display") & "'" 
			strSql = strSql & ", M_LINK2 = '" & ChkString(Request.Form("LINK2"),"display") & "'" 
		end if
		if strAge = "1" then
			strSql = strsql & ", M_AGE = '" & formbirthdate & "'"
		end if
		if strMarStatus = "1" then
			strSql = strSql & ", M_MARSTATUS = '" & ChkString(Request.Form("MarStatus"),"sqlstring") & "'" 
		end if
		if strSex = "1" then
			strSql = strsql & ", M_SEX = '" & ChkString(Request.Form("Sex"),"sqlstring") & "'" 
		end if
		if strOccupation = "1" then
			strSql = strSql & ", M_OCCUPATION='"& ChkString(Request.Form("Occupation"),"sqlstring") &"'" 
		end if
		if strBio = "1" then
			strSql = strSql & ", M_BIO = '" & ChkString(Request.Form("Bio"),"sqlstring") & "'" 
		end if
		if strHobbies = "1" then
			strSql = strSql & ", M_HOBBIES = '" & ChkString(Request.Form("Hobbies"),"sqlstring") & "'" 
		end if
		if strLNews = "1" then
			strsql = strsql & ", M_LNEWS = '" & ChkString(Request.Form("LNews"),"sqlstring") & "'" 
		end if
		if strQuote = "1" then
			strSql = strSql & ", M_QUOTE = '" & ChkString(Request.Form("Quote"),"sqlstring") & "'" 
		end if
		if strMSN = "1" then
			strSql = strSql & ", M_MSN = '" & ChkString(Request.Form("MSN"),"sqlstring") & "'"
		end if
		strSql = strSql & " WHERE MEMBER_ID = " & cLng(request.form("MEMBER_ID"))
		'response.Write(strSql & "<br>")	
		executeThis(strsql)
		
		if chkApp("forums","USERS") then			
		 if ChkString(Request.Form("Level"),"sqlstring") = "1" then 
		  ' - Remove the member from the moderator table
		  strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
		  strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & cLng(request.form("MEMBER_ID"))
		  executeThis(strsql)
		 end if
		end if		
	  	tmpResult = ""
	  	tmpResult = tmpResult & "<p align=center><span class=""fTitle"">" & txtProfUpd & "</span></p>"
		if sNewEmail= true and strEmail= 1 and (strEmailVal= 5 or strEmailVal= 6 or strEmailVal= 8) then
	  	  tmpResult = tmpResult & txtEmlHasChgd
		end if
	  	showUpdResult(tmpResult)
		%>
		<p align="center"><a href="members.asp"><%= txtBack %></a></p>
		<meta http-equiv="Refresh" content="3; URL=members.asp">
		<%
	  else
	  	tmpResult = ""
	  	tmpResult = tmpResult & "<p align=center><span class=""fTitle"">" & txtThereIsProb & "</span></p>"
	  	tmpResult = tmpResult & "<table align=center>"
	  	tmpResult = tmpResult & "<tr><td align=""center"">"
	  	tmpResult = tmpResult & "<ul>" & Err_Msg & "</ul>"
	  	tmpResult = tmpResult & "</td></tr></table>"
	    showUpdResult(tmpResult)
		 %>
		<p align=center><a href="JavaScript:history.go(-1)"><%= txtBack %></a></p>
	  <%
	  end if
	
	case else
	  tmpResult = ""
	  tmpResult = tmpResult & "<p align=center><span class=""fAlert""><b>" & txtERROR & "</b></span></p>"
	  tmpResult = tmpResult & txtNoPermViewPg
	  showUpdResult(tmpResult)
  end select
end sub

sub displayLCID()
strLCID_list = txtCountryLCID
arrLCID_list = split(strLCID_list,"|")

Response.Write("<select name=""intLCID"" id=""intLCID"">") & vbcrlf
xc = 0
for xl = 0 to ubound(arrLCID_list)
	if isLCID(split(arrLCID_list(xl),",")(1)) then
	xc = xc + 1
	  Response.Write("<option value=""" & split(arrLCID_list(xl),",")(1) & """" & chkSelect(intMemberLCID,cint(split(arrLCID_list(xl),",")(1))) & ">" & split(arrLCID_list(xl),",")(1) & " : " & split(arrLCID_list(xl),",")(0) &"</option>") & vbcrlf
  	  'response.Write( xc & ": <b>" & arrLCID_list(xl) & "</b><br>")
	else
  	  'response.Write(xc & ": " & arrLCID_list(xl) & "<br>")  intMemberLCID
	end if
next
Response.Write("</select>") & vbcrlf
end sub

Function isLCID(obj)
	on error resume next
	installed = False
	Err = 0
	session.LCID = obj
	If 0 = Err Then installed = True
	'Set chkObj = Nothing
	session.LCID = intMemberLCID
	isLCID = installed
	Err = 0
	on error goto 0
End Function
%>
