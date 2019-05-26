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

sub showMySubscriptions() 
  memID = strUserMemberID
	strSQL = "SELECT DISTINCT " & strTablePrefix & "SUBSCRIPTIONS.APP_ID, " & strTablePrefix & "APPS.APP_NAME, " & strTablePrefix & "APPS.APP_iNAME FROM " & strTablePrefix & "SUBSCRIPTIONS INNER JOIN " & strTablePrefix & "APPS ON " & strTablePrefix & "SUBSCRIPTIONS.APP_ID = " & strTablePrefix & "APPS.APP_ID WHERE (((" & strTablePrefix & "SUBSCRIPTIONS.M_ID)=" & strUserMemberID & ")) ORDER BY " & strTablePrefix & "APPS.APP_NAME;"
	set rsBmAp = my_Conn.execute(strSQL)
	If rsBmAp.Eof OR rsBmAp.Bof Then
	  showUpdResult(txtNoSubsFnd)
	Else
  	  do until rsBmAp.eof
  	  appID = rsBmAp("APP_ID")
      %>
        <form Action="cp_main.asp?cmd=6&mode=delete" method=post id="Form<%= appID %>" name="Form<%= appID %>">
	    <%
	    spThemeMM = "subscriptions"
	    spThemeTitle = ucase(rsBmAp("APP_NAME")) & " " & txtSubsc
	    spThemeBlock1_open("1")
	  sSQL = "SELECT " & strTablePrefix & "SUBSCRIPTIONS.SUBSCRIPTION_ID, " & strTablePrefix & "SUBSCRIPTIONS.APP_ID, " & strTablePrefix & "SUBSCRIPTIONS.ITEM_ID, " & strTablePrefix & "SUBSCRIPTIONS.SUBCAT_ID, " & strTablePrefix & "SUBSCRIPTIONS.CAT_ID, " & strTablePrefix & "SUBSCRIPTIONS.ITEM_TITLE, " & strTablePrefix & "APPS.APP_iNAME, " & strTablePrefix & "APPS.APP_VIEW, " & strTablePrefix & "SUBSCRIPTIONS.M_ID FROM " & strTablePrefix & "SUBSCRIPTIONS INNER JOIN " & strTablePrefix & "APPS ON " & strTablePrefix & "SUBSCRIPTIONS.APP_ID = " & strTablePrefix & "APPS.APP_ID WHERE (((" & strTablePrefix & "SUBSCRIPTIONS.APP_ID)=" & appID & ") AND ((" & strTablePrefix & "SUBSCRIPTIONS.M_ID)=" & memID & ")) ORDER BY " & strTablePrefix & "SUBSCRIPTIONS.CAT_ID DESC, " & strTablePrefix & "SUBSCRIPTIONS.SUBCAT_ID DESC, " & strTablePrefix & "SUBSCRIPTIONS.ITEM_ID DESC;"
	  set rsBmks = my_Conn.execute(sSQL)
	  'response.Write("app iname: " & rsBmks("APP_iNAME") & ":<br>")
	    curType = 0
		shoHeader = false
	    do while not rsBmks.eof
		  select case rsBmks("APP_iNAME")
		    case "forums"
			  if rsBmks("ITEM_ID") = 0 then
			    if curtype <> 1 then
				  shoHeader = true
				end if
				subType = txtForums & ":"
				curType = 1
			    lnkTo = "link.asp?forum_id=" & rsBmks("SubCat_ID")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  else
			    if curtype <> 2 and not shoHeader then
				  shoHeader = true
				end if
				subType = txtTopics & ":"
				curType = 2
			    lnkTo = "link.asp?topic_id=" & rsBmks("ITEM_ID")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  end if
		    case "links"
			  if rsBmks("SUBCAT_ID") = 0 and rsBmks("CAT_ID") = 0 then 'cat bookmark
			    lnkTo = rsBmks("APP_VIEW")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  elseif rsBmks("SUBCAT_ID") = 0 and rsBmks("CAT_ID") <> 0 then 'subcat subscription
			    if curtype <> 1 then
				  shoHeader = true
				end if
				itmTitle = rsBmks("ITEM_TITLE")
				subType = txtCats & ":"
				curType = 1
			    lnkTo = "links.asp?cmd=1&cid=" & rsBmks("CAT_ID")
				cls = "tCellAlt1"
			  else
			    if curtype <> 2 and not shoHeader then
				  shoHeader = true
				end if
				itmTitle = rsBmks("ITEM_TITLE")
				sSQL = "SELECT LINKS_SUBCATEGORIES.SUBCAT_NAME, LINKS_CATEGORIES.CAT_NAME FROM LINKS_SUBCATEGORIES INNER JOIN LINKS_CATEGORIES ON LINKS_SUBCATEGORIES.CAT_ID = LINKS_CATEGORIES.CAT_ID WHERE (((LINKS_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				set rsB = my_Conn.execute(sSQL)
				if not rsB.eof then
				    itmTitle = rsB(1) & "/" & rsB(0)
				end if
				set rsB = nothing
				subType = txtSubCats & ":"
				curType = 2
			    lnkTo = "links.asp?cmd=2&sid=" & rsBmks("SUBCAT_ID")
				cls = "tCellAlt1"
			  end if
		    case else
			  if rsBmks("SUBCAT_ID") = 0 and rsBmks("CAT_ID") = 0 then 'cat bookmark
			    lnkTo = rsBmks("APP_VIEW")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  elseif rsBmks("SUBCAT_ID") = 0 and rsBmks("CAT_ID") <> 0 then 'subcat subscription
			    if curtype <> 1 then
				  shoHeader = true
				end if
				itmTitle = rsBmks("ITEM_TITLE")
				subType = txtCats & ":"
				curType = 1
			    lnkTo = rsBmks("APP_VIEW") & "?cmd=1&cid=" & rsBmks("CAT_ID")
				cls = "tCellAlt1"
			  else
			    if curtype <> 2 and not shoHeader then
				  shoHeader = true
				end if
				itmTitle = rsBmks("ITEM_TITLE")
				select case rsBmks("APP_iNAME")
				  case "article"
					sSQL = "SELECT ARTICLE_SUBCATEGORIES.SUBCAT_NAME, ARTICLE_CATEGORIES.CAT_NAME FROM ARTICLE_SUBCATEGORIES INNER JOIN ARTICLE_CATEGORIES ON ARTICLE_SUBCATEGORIES.CAT_ID = ARTICLE_CATEGORIES.CAT_ID WHERE (((ARTICLE_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				  case "pictures"
					sSQL = "SELECT PIC_SUBCATEGORIES.SUBCAT_NAME, PIC_CATEGORIES.CAT_NAME FROM PIC_SUBCATEGORIES INNER JOIN PIC_CATEGORIES ON PIC_SUBCATEGORIES.CAT_ID = PIC_CATEGORIES.CAT_ID WHERE (((PIC_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				  case "downloads"
					sSQL = "SELECT DL_SUBCATEGORIES.SUBCAT_NAME, DL_CATEGORIES.CAT_NAME FROM DL_SUBCATEGORIES INNER JOIN DL_CATEGORIES ON DL_SUBCATEGORIES.CAT_ID = DL_CATEGORIES.CAT_ID WHERE (((DL_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				  case "classifieds"
					sSQL = "SELECT CLASSIFIED_SUBCATEGORIES.SUBCAT_NAME, CLASSIFIED_CATEGORIES.CAT_NAME FROM CLASSIFIED_SUBCATEGORIES INNER JOIN CLASSIFIED_CATEGORIES ON CLASSIFIED_SUBCATEGORIES.CAT_ID = CLASSIFIED_CATEGORIES.CAT_ID WHERE (((CLASSIFIED_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				end select
				set rsB = my_Conn.execute(sSQL)
				if not rsB.eof then
				    itmTitle = rsB(1) & "/" & rsB(0)
				end if
				set rsB = nothing
				subType = txtSubCats & ":"
				curType = 2
			    lnkTo = rsBmks("APP_VIEW") & "?cmd=2&sid=" & rsBmks("SUBCAT_ID")
				cls = "tCellAlt1"
			  end if
		  end select
		  if shoHeader then
		    shoHeader = false
	       response.Write("<div class=""tAltSubTitle"" style=""padding-left:8px;padding-top:5px;padding-bottom:3px; text-align:left;"">&nbsp;&nbsp;<b>" & subType & "</b></div>")
			
		  end if
	      response.Write("<div class=""" & cls & """ style=""padding-left:8px;padding-top:5px;padding-bottom:3px; text-align:left;""><input type=""checkbox"" name=""delBookmark"" value=""" & rsBmks("SUBSCRIPTION_ID") & """>&nbsp;&nbsp;<b><a href=""" & lnkTo & """>" & itmTitle & "</a></b></div>")
	      rsBmks.movenext
	    loop
	 %>
	  <div align="center">
	  <input type="submit" name="del" value="<%= replace(txtDelSelSubsc,"[%name%]",ucase(rsBmAp("APP_NAME"))) %>" ID="Submit<%= rsBmAp("APP_ID") %>" class="button"></div><%
	    spThemeBlock1_close("1") %>
        </FORM><br />
<%    rsBmAp.movenext
      loop
    end if
	set rsBmAp = nothing
	set rsBmks = nothing
end sub

sub modifySubscriptions()
Select Case Request.QueryString("mode")

Case "delete"
	delCnt = 0
	delBkmk = split(Request.Form("delBookmark"), ",")
	
	for ib = 0 to ubound(delBkmk)
		if trim(delBkmk(ib)) <> "" then
		chkString delBkmk(ib),"numeric"
		' Delete selected topic bookmarks
		delSQL = "DELETE FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE SUBSCRIPTION_ID = " & delBkmk(ib)
    	executeThis(delSQL)
		delCnt = delCnt + 1
		end if
	next
	if delCnt > 0 then
	  tmpResult = tmpResult & txtSelSubDel
	else
	  tmpResult = tmpResult & txtSelSubToDel
	end if
  	'showUpdResult(tmpResult)

Case "deleteAll"
	'delBookmark = split(Request.Form("delBookmark"), ",")
		delSQL = "DELETE FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID = " & memID
    executeThis(delSQL)
	
	  tmpResult = tmpResult & txtAllSubsDel
  	  'showUpdResult(tmpResult)
  end select
end sub
%>