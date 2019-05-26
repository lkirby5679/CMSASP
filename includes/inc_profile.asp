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
 on error resume next 
 %> 

	<table border="0" width="100%" cellspacing="0" cellpadding="0" valign="top">
	  <tr>
	    <td align=left valign="top">
<%
'spThemeBlock1_open(intSkin)
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td align="center" colspan="2" class="tCellAlt0">
		<p><b><%= txtReg1a %>&nbsp;<span class="fAlert"><b>*</b></span>&nbsp;<%= txtReg1b %></b>
<%		if lcase(strEmail) = "1" And (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8) then
			If Request.Querystring("mode") = "Register" Then %>
				<br><span class="fAlert"><%= txtReg2a %></span>.</p><p><%= txtReg3a %>&nbsp;<%= replace(strSender,"@","[no-spam]@") %>.<br><%= txtReg3b %>&nbsp;"[no-spam]"<br><%= txtReg3c %>.<br><br></p>
<%			else %>
				<br><%= txtReg2b %>.</p><p><%= txtReg3a %>&nbsp;<%= replace(strSender,"@","[no-spam]@") %>.<br><%= txtReg3b %>&nbsp;"[no-spam]"<br><%= txtReg3c %>.<br><br></p>
<%      		
			end if
		end if%><!-- S k y D o g g - S k y P o r t a l - is here - december 2005-->
	    </td>
	  </tr>
  <tr>
	<td colspan="2" class="tCellAlt0" valign="top">
      <table border="0" width="100%" cellspacing="0" cellpadding="1">
<%
'<!-- ::::::::::::::::::::::::: start BASICS info ::::::::::::::::::::::::::::: --> %>
        <tr> 
          <td valign="top" align=center colspan="2" class="tSubTitle"><b><%= txtBasics %></b></td>
        </tr>
	    <%
        if Request.Querystring("mode")="goModify" or Request.Querystring("mode")="goEdit" then %>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="center"><br></td>
        </tr>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="center">
		  <b><%= txtRefFrndUrl %>: </b></td>
        </tr>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="center">  
            <%= strHomeURL %>policy.asp?rname=<%=rs("M_NAME")%><br>&nbsp;</td>
        </tr>
        <%
		end if %>
        <% if not Request.Querystring("mode") = "goEdit" then %>
        <%   if (trim(Request.QueryString("rname") = "")) then %>
        <tr> 
          <td class="tCellAlt0" width="40%" align="right" nowrap><b><%= txtRefer %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT name="Referrer" size="25" maxlength="90" value="<%= ChkString(RS("M_RNAME"), "display") %>">
          </td>
        </tr>
        <%   else %>
        <tr> 
          <td class="tCellAlt0" width="40%" align="right" nowrap><b><%= txtRefer %>:&nbsp;</b></td>
          <td class="tCellAlt0" align=left nowrap> 
            <%= ChkString(Request.Querystring("rname"), "sqlstring") %>
            </td>
          <INPUT type="hidden" name="Referrer" value="<%= ChkString(Request.Querystring("rname"),"sqlstring") %>">
        </tr>
        <%   end if 
		   end if%>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap width="40%">
		  <b><span class="fAlert">*</span><%= txtUsrNam %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
      <%if (Request.QueryString("mode") = "goEdit" or Request.Querystring("mode") = "goModify") and mLev < 4 then %>
            <% =ChkString(rs("M_NAME"), "display") %>
            <INPUT type="hidden" name="Name"  value="<% =rs("M_NAME") %>">
      <%else %>
            <INPUT name="Name" size="25" maxlength="90"  value="<% =ChkString(rs("M_NAME"), "display") %>">
      <%end if %>
            </td>
        </tr>
        <%
		if strAuthType = "nt" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top">
		  <b><span class="fAlert">*</span><%= txtUAcct %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <%if Request.Querystring("mode") = "goModify" and mlev = 4 then %>
            <input name="Account" value="<% =ChkString(rs("M_USERNAME"), "display") %>" size="20">
            <%else %>
            <%=Session(strCookieURL & "userid")%> 
            <input type="hidden" name="Account" value="<% =Session(strCookieURL & "userid") %>">
            <%end if %>
            </td>
        </tr>
        <%
		else %>
        <tr> 
            <%
			  ast = "*"
			  stTxt = txtPass
			  if Request.QueryString("mode") <> "Register" then
			    ast = ""%>
            <% stTxt = txtNew & " " & stTxt %> 
            <%end if%>
          <td class="tCellAlt0" align="right" nowrap><b><span class="fAlert"><%= ast %></span> 
            <%= stTxt %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT name="Password" type="Password" size="25" maxlength="90" value="">
            <INPUT name="Password-d" type=hidden value="<% =rs("M_PASSWORD") %>">
            </td>
        </tr>
        <%	if Request.QueryString("mode") = "Register" or Request.QueryString("mode") = "goEdit" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><span class="fAlert"><%= ast %></span> 
            <%= txtPassAgn %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT name="Password2" type="Password" value="" size="25">
            </td>
        </tr>
        <%	end if
		end if 
		if mlev = 4 and Request.Querystring("mode") = "Register" then %>
        <tr> 
          <td align="right" nowrap>
		  <b><%= txtResUNam %>:</b>&nbsp;</td>
          <td class="tCellAlt0">
		  <input type="checkbox" Value="yes" name="reservation">
          </td>
        </tr>
        <tr> 
          <td align="right" nowrap><b><%= txtEmlNewUsr %>:</b>&nbsp;</td>
          <td class="tCellAlt0"><input type="checkbox" Value="yes" name="sendinvite">
          </td>
        </tr>
        <%
		end if
		if Request.Querystring("mode")="goModify" and mLev = 4 then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtMlev %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <SELECT value="1" name="Level">
              <OPTION VALUE="1"<% if rs("M_LEVEL") = 1 then Response.Write(" selected") %>><%= txtNmlUsr %> 
              <OPTION VALUE="2"<% if rs("M_LEVEL") = 2 then Response.Write(" selected") %>><%= txtModerator %> 
              <OPTION VALUE="3"<% if rs("M_LEVEL") = 3 then Response.Write(" selected") %>><%= txtAdminst %> 
            </SELECT>
          </td>
        </tr>
        <%
		end if
		if Request.Querystring("mode") = "goModify" and mLev = 4 then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtTitle %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="Title" size="25" maxlength="90" value="<%= CleanCode(RS("M_TITLE")) %>">
            </td>
        </tr>
        <%
		end if
		if strFullName = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtFNam %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <input name="FirstName" value="<% =rs("M_FIRSTNAME") %>" size="20">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtLNam %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <input name="LastName" value="<% =rs("M_LASTNAME") %>" size="20">
            </td>
        </tr>
        <%
		end if
		if strCity = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtCity %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <input name="City" value="<% =rs("M_CITY") %>" size="20">
            </td>
        </tr>
        <%
		end if
		if strState = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtState %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <input name="State" value="<% =rs("M_STATE") %>" size="20">
            </td>
        </tr>
        <%
		end if
		if strZip = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtZipCd %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <input name="Zipcode" value="<% =rs("M_ZIP") %>" size="20">
            </td>
        </tr>
        <%
		end if
		if strCountry = "1" Then
           strSQL="Select CO_NAME from "& strTablePrefix & "COUNTRIES ORDER by CO_NAME"
           Set rstCO = my_Conn.Execute (strSql) %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtCntry %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <select name="Country" size="1">
            <%If Request.Querystring("mode") = "Register" Then %>
                 <option value="" selected></option>
            <%Else %>
                 <option value="" ></option>
            <%End If 
		  rstCO.movefirst
          While (Not rstCO.EOF) %>
              <option value="<%=(rstCO.Fields.Item("CO_NAME").Value)%>"
               <%If (Request.Querystring("mode") <> "Register") Then
                    If rstCO("CO_NAME")= trim(rs("M_COUNTRY")) Then
                    Response.Write(" SELECTED")
                    End if
                 Else
                  Response.Write("")
                 End If %>><%=ChkString(rstCO("CO_NAME"), "display")%></option>
         <%
            rstCO.MoveNext()
          Wend
          If (rstCO.CursorType > 0) Then
          rstCO.MoveFirst
          Else
          rstCO.Requery
          End If %>
                </select>
            </td>
        </tr>
        <%
		  Set rsCO = Nothing
        end If
		if strAge = "1" Then %>
		<tr>
		  <td class="tCellAlt0" align="right" nowrap valign="top">
			<b><%= txtBDate %>:&nbsp;</b></td>
		  <td class="tCellAlt0">
			<select name="B_Month" size="1">
			<% If rs("M_AGE") = "" or rs("M_AGE") = " " Then %>
            <option value=" " selected></option>
			  <% for i = 1 to 12 
                   if i < 10 then
                     iStr = "0" + Cstr(i)
                   else 
                     iStr = Cstr(i)
                   end if %>
                   <option value="<%=iStr%>"><% =Monthname(i)%></option>
              <% next %>
			<% Else %>
                 <option value=" "></option>
				 <% for i = 1 to 12 
                      if i < 10 then
                        iStr = "0" + Cstr(i)
                      else 
                        iStr = Cstr(i)
                      end if
                      if left(rs("M_AGE"),2) = iStr then %>
                        <option value="<%=iStr%>" selected><% =Monthname(i) %></option>
                   <% else %>
                         <option value="<%=iStr%>"><% =Monthname(i)%></option>
                   <% end if
                   next
			   End If %>
               </select>
       		   <select name="B_Day" size="1">
			   <% If rs("M_AGE") = "" or rs("M_AGE") = " " Then %>
                    <option value=" " selected></option>     
                    <% for i = 1 to 31 
                         if i < 10 then
                           iStr = "0" + Cstr(i)
                         else 
                           iStr = Cstr(i)
                         end if %>
                         <option><%=iStr%></option>
                    <% next %>
			    <% Else %>              
                    <option value=" "></option>                 
                    <% for i = 1 to 31 
                         if i < 10 then
                           iStr = "0" + Cstr(i)
                         else 
                           iStr = Cstr(i)
                         end if %>
                      <% if mid(rs("M_AGE"),4,2) = iStr then %>
                           <option selected><%=iStr%></option>
                      <% else %>
                                <option><%=iStr%></option>
                      <% end if %>
                    <% next %>
				<% End If %>
                </select>
				<SELECT NAME="B_YEAR">
				<% If rs("M_AGE") = "" or rs("M_AGE") = " " Then %>
                     <option value=" " selected></option>   
                     <%for i = -100 to 0%>
                      <% dtToday = date()
	                     intThisYear = Year(dtToday) %>
	                     <OPTION VALUE=<%= intThisYear + i%>><%= intThisYear + i%></option>
                     <%next%>
				<% Else %>              
                     <option value=" "></option>     
                     <%for i = -100 to 0
                         dtToday = date()
	                     intThisYear = Year(dtToday) %>
	                     <OPTION VALUE=<%= intThisYear + i%> <%if Cstr(intThisYear + i) = mid(rs("M_AGE"),7,4) then Response.Write(" selected")%>><%= intThisYear + i%></option>
                     <%next%>
				<% End If %>
                </SELECT>
				</td>
			  </tr>
        <%
		end if
		if strSex = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtSex %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <select name="Sex" size="1">
			<% If RS("M_SEX") = "" or RS("M_SEX") = " " Then %>
              <OPTION VALUE=" " selected></option>
              <OPTION VALUE="Male">&nbsp;<%= txtM %>&nbsp;</option>
              <OPTION VALUE="Female">&nbsp;<%= txtF %>&nbsp;</option>
			<% Else %>
			  	<% If RS("M_SEX") = "Male" Then %>
              	<OPTION VALUE=" "></option>
              	<OPTION VALUE="Male" selected>&nbsp;<%= txtM %>&nbsp;</option>
              	<OPTION VALUE="Female">&nbsp;<%= txtF %>&nbsp;</option>
				<% Else %>
              	<OPTION VALUE=" "></option>
              	<OPTION VALUE="Male">&nbsp;<%= txtM %>&nbsp;</option>
              	<OPTION VALUE="Female" selected>&nbsp;<%= txtF %>&nbsp;</option>
				<% End If %>
			<% End If %>
            </select>
            </td>
        </tr>
        <%
		end if
		if strMarStatus = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top">
		  <b><%= txtMarStat %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="MarStatus" size="25" maxlength="90" value="<% =ChkString(RS("M_MARSTATUS"), "display") %>">
            </td>
        </tr>
        <%
		end if
		if strOccupation = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtOcc %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="Occupation" size="25" maxlength="90" value="<% = ChkString(RS("M_OCCUPATION"), "display") %>">
            </td>
        </tr>
        <%
		end if
		if Request.Querystring("mode") = "goModify" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap>
		  <b><%= txtPosts %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="Posts" size="25" maxlength="90" value="<% = ChkString(RS("M_POSTS"), "display") %>">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtGold %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="Gold" size="25" maxlength="90" value="<% = ChkString(RS("M_GOLD"), "display") %>">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap>
		  <b><%= txtRepPts %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="Rep" size="25" maxlength="90" value="<% = ChkString(RS("M_REP"), "display") %>">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap>
		  <b><%= txtRfls %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <INPUT  name="Referrals" size="25" maxlength="90" value="<% = ChkString(RS("M_RTOTAL"), "display") %>">
            </td>
        </tr>
        <%
		end if %>
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<%
'<!-- ::::::::::::::::::::::::: END BASICS info ::::::::::::::::::::::::::::: -->

'<!-- ::::::::::::::::::::::::: start contact info ::::::::::::::::::::::::::::: -->
%>
        <tr> 
          <td align=center class="tSubTitle" colspan="2"><b>&nbsp;<%= txtCtInfo %>&nbsp;</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" width="40%" nowrap><span class="fAlert"><b>*</b></span><b><%= txtEmlAdd %>:&nbsp;</b></td>
          <td class="tCellAlt0">
		  <% if CurPageType = "register" then
		  	   sEmail = ""
		     else
		  	   sEmail = ChkString(RS("M_EMAIL"), "display")
		     end if %>
            <INPUT type="textbox" name="Email" size="25" maxlength="90" value="<%= sEmail %>"><INPUT type="hidden" name="Email3" value="<%= sEmail %>">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><span class="fAlert">*</span><%= txtCfmEml %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="Email2" size="25" maxlength="90" value="<%= sEmail %>">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtRecEml %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
        <%if Request.QueryString("mode") = "goEdit" or Request.Querystring("mode") = "goModify" then 
			RECMAIL = RS("M_RECMAIL")
%>
            <input type="radio" Value="0" name="recmail" <%if RECMAIL = "0" then%>checked<%end if%>> <%= txtYes %>
            <input type="radio" Value="1" name="recmail" <%if RECMAIL = "1" then%>checked<%end if%>> <%= txtNo %> 
	    <%else %>
            <input type="radio" Value="0" name="recmail" checked>
            <%= txtYes %>
            <input type="radio" Value="1" name="recmail">
            <%= txtNo %> 
		<%end if %>
		  </td>
        </tr>
        <% 'end if %>
        <tr valign="center"> 
          <td class="tCellAlt0" align="right" valign="center" width="10%" nowrap>
		  <b><%= txtHidEml %>:</b></td>
          <% if Request.QueryString("mode") = "Register" then %>
          <td class="tCellAlt0" valign="center">
            <input type="radio" name="HideMail" value="1">
            <%= txtYes %>
            <input type="radio" name="HideMail" value="0" checked>
            <%= txtNo %></td>
          <% else %>
          <td class="tCellAlt0" valign="center">
            <input type="radio" name="HideMail" value="1"<% if RS("M_HIDE_EMAIL") <> "0" then Response.Write("checked") %>>
            <%= txtYes %>
            <input type="radio" name="HideMail" value="0"<% if RS("M_HIDE_EMAIL") = "0" then Response.Write("checked") %>>
            <%= txtNo %></td>
          <% end if %>
        </tr>
        <%
		if strMSN = "1" then %>
        <tr>
          <td class="tCellAlt0" align="right" nowrap><b><%= txtMSN %>:&nbsp;</b></td>
          <td class="tCellAlt0">
		  <% 
		  if trim(ChkString(RS("M_MSN"), "display")) = "" then
		  	stringMSN = " "
		  else 
		  	stringMSN = trim(ChkString(RS("M_MSN"), "display"))
		  end if %>
            <INPUT name="MSN" size="25" maxlength="90" value="<% =stringMSN %>">
          </td>
        </tr>
        <% 
		end if
		if strICQ = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtICQ %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="ICQ" size="25" maxlength="90" value="<% =ChkString(RS("M_ICQ"), "display") %>">
            </td>
        </tr>
        <%
		end if
		if strYAHOO = "1" then
		%>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtYhoIM %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="YAHOO" size="25" maxlength="90" value="<% =ChkString(RS("M_YAHOO"), "display") %>">
            </td>
        </tr>
        <%
		end if
		if strAIM = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap><b><%= txtAIM %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="AIM" size="25" maxlength="90" value="<% =ChkString(RS("M_AIM"), "display") %>">
            </td>
        </tr>
        <%
		end if %>
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<%
'<!-- ::::::::::::::::::::::::: END contact info ::::::::::::::::::::::::::::: -->

'<!-- ::::::::::::::::::::::::: start PHOTO info ::::::::::::::::::::::::::::: -->
        if strPicture = "1" then %>
        <tr> 
          <td align=center class="tSubTitle" colspan="2"> <b><%= txtMPic %>&nbsp;</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap valign="top"><b><%= txtMPicUrl %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT  name="Photo_URL" size="25" maxlength="90" value="<%  if Trim(RS("M_PHOTO_URL") <> "") and lcase(RS("M_PHOTO_URL")) <> "http://" and not(IsNull(RS("M_PHOTO_URL"))) then Response.Write(ChkString(rs("M_PHOTO_URL"), "display")) else Response.Write("http://") %>">
            </td>
        </tr>
        <%
		end if%>
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<% ' strPicture
'<!-- ::::::::::::::::::::::::: END PHOTO info ::::::::::::::::::::::::::::: -->

'<!-- ::::::::::::::::::::::::: start AVATAR info ::::::::::::::::::::::::::::: -->
		 %>
        <!--#INCLUDE file="inc_avatar.asp" -->
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<%
'<!-- ::::::::::::::::::::::::: END AVATAR info ::::::::::::::::::::::::::::: -->

'<!-- ::::::::::::::::::::::::: start LINKS info ::::::::::::::::::::::::::::: -->

		if (strHomepage + strFavLinks) > 0 then %>
        <tr> 
          <td align=center class="tSubTitle" colspan="2"> <b><%= txtLinks %>&nbsp;</b></td>
        </tr>
        <%	if strHomepage = "1" then %>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap width="10%"><b><%= txtHmPg %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="Homepage" size="25" maxlength="90" value="<% if ChkString(RS("M_Homepage"), "display") <> " " and lcase(RS("M_Homepage")) <> "http://" then Response.Write(ChkString(RS("M_Homepage"),"display")) else Response.Write("http://") %>">
            </td>
        </tr>
        <%	end if
			if strFavLinks = "1" then%>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap width="10%"><b><%= txtColLnks %>:&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="Link1" size="25" maxlength="90" value="<% if RS("M_LINK1") <> " " and lcase(RS("M_LINK1")) <> "http://" then Response.Write(ChkString(rs("M_LINK1"), "display")) else Response.Write("http://") %>">
            </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" align="right" nowrap width="10%"><b>&nbsp;</b></td>
          <td class="tCellAlt0">
            <INPUT name="Link2" size="25" maxlength="90" value="<% if RS("M_LINK2") <> " " and lcase(RS("M_LINK2")) <> "http://" then Response.Write(ChkString(rs("M_LINK2"), "display")) else Response.Write("http://") %>">
            </td>
        </tr>
        <%	end if
		end if %>
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<%
'<!-- ::::::::::::::::::::::::: END LINKS info ::::::::::::::::::::::::::::: -->

'<!-- ::::::::::::::::::::::::: start MORE ABOUT ME info ::::::::::::::::::::::::::::: -->
		if (strBio + strHobbies + strLNews + strQuote) > 0 then 
				strMyBio = rs("M_BIO")
				strMyHobbies = rs("M_HOBBIES")
				strMyLNews = rs("M_LNEWS")
				strMyQuote = rs("M_QUOTE") %>
        <tr> 
          <td align=center class="tSubTitle" colspan="2">
		  <b><%= txtMAbtMe %></b></td>
        </tr>
        <% if strBio = "1" then  %>
        <tr> 
          <td class="tCellAlt0" valign="top" align="right" nowrap width="10%"> <b><% =strVar1 %>:&nbsp;</b> 
          </td>
          <td class="tCellAlt0" valign="top"> 
            <textarea name="Bio" cols="30" rows=4><% =Trim(cleancode(strMyBio)) %></textarea>
             </td>
        </tr>
        <% end if
		   if strHobbies = "1" then %>
        <tr> 
          <td class="tCellAlt0" valign="top" align="right" nowrap width="10%"> <b><% =strVar2 %>:&nbsp;</b> 
          </td>
          <td class="tCellAlt0" valign="top"> 
            <textarea name="Hobbies" cols="30" rows=4><% =Trim(cleancode(strMyHobbies)) %></textarea>
             </td>
        </tr>
        <% end if
		   if strLNews = "1" then  %>
        <tr> 
          <td class="tCellAlt0" valign="top" align="right" nowrap width="10%"> <b><% =strVar3 %>:&nbsp;</b> </td>
          <td class="tCellAlt0" valign="top"> 
            <textarea name="LNews" cols="30" rows=4><% =Trim(cleancode(strMyLNews)) %></textarea>
             </td>
        </tr>
        <% end if
		   if strQuote = "1" then %>
        <tr> 
          <td class="tCellAlt0" valign="top" align="right" nowrap width="10%"> <b><% =strVar4 %>:&nbsp;</b></td>
          <td class="tCellAlt0" valign="top"> 
            <textarea name="Quote" cols="30" rows=4><% =Trim(cleancode(strMyQuote)) %></textarea>
             </td>
        </tr>
        <% end if
		end if %>
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<%
'<!-- ::::::::::::::::::::::::: END MORE ABOUT ME info ::::::::::::::::::::::::::::: -->
		%>
        <tr> 
          <td colspan="2" align="center" class="tSubTitle"><%= txtSigatr %></td>
        </tr>
        <tr> 
          <td align=center colspan="2">&nbsp;</td>
        </tr>
		<%
		If strAllowHtml = 1 and hasEditor Then
				strTxtSig = Trim(RS("M_SIG")) %>
        <tr>
          <td class="tCellAlt0" colspan="2" align="center"> 
            <textarea maxLength="255" name="Sig" cols="50" rows="15"><% =strTxtSig %></textarea>
          </td>
        </tr>
  		<%
		else
		  strTxtSig = Trim(cleancode(RS("M_SIG"))) %>
        <tr> 
          <td class="tCellAlt0" align="right" valign="top" nowrap><b><%= txtSigatr %>:&nbsp;</b></td>
          <td class="tCellAlt0"> 
            <textarea maxLength="255" name="Sig" cols="25" rows=4><% =strTxtSig %></textarea>
          </td>
        </tr>
        <tr> 
          <td class="tCellAlt0" colspan="2" align="right" valign="top" nowrap> 
            <center>
              <input name="Preview" type="button" value=" <%= txtPrevSig %> " class="Button" onclick="OpenSPreview()">
            </center>
          </td>
        </tr>
  		<%				
  		end if %>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="right" nowrap>&nbsp;</td>
		</tr>
        <%
        if Request.Querystring("mode")="goModify" or Request.Querystring("mode")="goEdit" then %>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="center"><br></td>
        </tr>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="center">
		  <b><%= txtRefFrndUrl %>: </b></td>
        </tr>
        <tr> 
          <td colspan="2" class="tCellAlt0" align="center">  
            <%= strHomeURL %>policy.asp?rname=<%=rs("M_NAME")%></td>
        </tr>
        <%
		end if %>
      </table>
	  </td>
	</tr>
	<tr><td class="tCellAlt0" nowrap align="center" valign="middle" colspan="2">&nbsp;
<%  If SecImage >0 And Request.Querystring("mode") = "Register" Then %>
	<br><%= txtEntrSecImg %>
	<input CLASS="textbox" type="text" name="secCode" size="8" maxLength="8" value="" onFocus="javascript:this.value='';">
&nbsp;&nbsp;<img align="absolute" src="<%= strHomeUrl %>includes/securelog/image.asp" />
<%  end if %>
	</td>
	</tr>
	<tr><td class="tCellAlt0" nowrap align="center" valign="middle" colspan="2">
        <INPUT type="hidden" value="<%= chkString(Request.Form("MEMBER_ID"),"sqlstring") %>" name="MEMBER_ID">
        <INPUT type="submit" value="  <%= txtSubmit %>  " name="Submit1" class="button"><br><br>
	</td>
	</tr>
	</table>
<%
'spThemeBlock1_close(intSkin)%>
	</td>
  </tr> 
</table>