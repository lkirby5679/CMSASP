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

function chkLoginStatus()
  if (strAuthType = "nt") then
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
	strDBNTUserName = Session(strCookieURL & "userID")
	strDBNTFUserName = Session(strCookieURL & "userID")
  end if

  if strAuthType = "db" then
	if (Request.Cookies(strUniqueID & "User")("Name") <> "" and Request.Cookies(strUniqueID & "User")("PWord") <> "") then
		'
		strSql = "SELECT MEMBER_ID, M_NAME, M_LEVEL, M_EMAIL, M_PASSWORD, M_PMSTATUS, M_PMRECEIVE"
		strSql = strSql & ", M_TIME_OFFSET, M_TIME_TYPE, M_LCID, M_AGE"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE M_NAME = '" & ChkString(Request.Cookies(strUniqueID & "User")("Name"), "SQLString") & "' "
		strSql = strSql & " AND M_PASSWORD = '" & ChkString(Request.Cookies(strUniqueID & "User")("Pword"), "SQLString") &"' and M_STATUS=1"
		'response.Write(strSql & "<br><br>")
		Set rsCheck = my_Conn.Execute(strSql)
		if rsCheck.BOF and rsCheck.EOF then
			Call ClearCookies()
			strDBNTUserName = ""
			strUserMemberID = 0
			strUserEmail = ""
			mLev = 0
			PMaccess = 0
		else
			strDBNTUserName = rsCheck("M_NAME")
			strUserMemberID = clng(rsCheck("MEMBER_ID"))
			strUserEmail = rsCheck("M_EMAIL")
			mLev = rsCheck("M_LEVEL")+1
			if mLev = 4 then
			  intIsSuperAdmin = chkIsSuperAdmin(2,strDBNTUsername)
			end if
			strMBirthday = rsCheck("M_AGE")
			strMTimeAdjust = rsCheck("M_TIME_OFFSET")
			strMTimeType = rsCheck("M_TIME_TYPE")
			intMemberLCID = rsCheck("M_LCID")
			if len(intMemberLCID) = 4 or len(intMemberLCID) = 5 then
			  Session.LCID = intMemberLCID
			end if
			
			strTimeType = strMTimeType
			strMCurDateAdjust = DateAdd("h", (strTimeAdjust + strMTimeAdjust) , now())
			strMCurDateString = DateToStr2(strMCurDateAdjust)
			strForumDateAdjust = ChkDate2(strMCurDateString)
			strForumTimeAdjust = strForumDateAdjust & chkTime2(strMCurDateString)
			
			PMaccess = rsCheck("M_PMSTATUS")
			if rsCheck("M_PMRECEIVE") = 0 then
			  PMaccess = 0
			end if
  			' checks and sets the members PM count
  			pmCheck()
		end if
		set rsCheck = nothing
	else
		strDBNTUserName = ""
		strUserMemberID = 0
		strUserEmail = ""
		mLev = 0
		PMaccess = 0
	end if
  end if
end function

function bldArrUserGroup()
  if mLev > 0 then ':: they are a member
	strSql = "SELECT G_GROUP_ID, G_GROUP_LEADER FROM " & strTablePrefix & "GROUP_MEMBERS WHERE G_MEMBER_ID = " & strUserMemberID
	set rsApp = my_Conn.execute(strSql)
	if not rsApp.eof then
		tmpArr1 = "2," 'add member group by default
		tmpArr2 = "0,"
		do until rsApp.eof
			tmpArr1 = tmpArr1 & rsApp("G_GROUP_ID") & ","
			tmpArr2 = tmpArr2 & rsApp("G_GROUP_LEADER") & ","
			rsApp.movenext
		loop
		if tmpArr1 <> "" then
			tmpArr3 = split(tmpArr1,",")
			tmpArr4 = split(tmpArr2,",")
			acnt = ubound(tmpArr3)-1
			redim arrGroups(acnt,1)
			for ag = 0 to ubound(tmpArr3)-1
				arrGroups(ag,0) = tmpArr3(ag)
				arrGroups(ag,1) = tmpArr4(ag)
			next
		end if
	else
		redim arrGroups(0,1)
	 	arrGroups(0,0) = "2" 'members group
	 	arrGroups(0,1) = "0" 'not group leader
	end if
	set rsApp = nothing
	
  else '::they are a guest
	redim arrGroups(0,1)
	arrGroups(0,0) = "3" 'EVERYONE group
	arrGroups(0,1) = "0" 'not group leader
  end if
end function

function bldArrAppAccess()
	dim tmpAppID, tmpAppActive, tmpAppGroupsR, tmpAppGroupsW, tmpAppGroupsF, tmpAppIName, bHasAccess
	dim tmpAppSubsc, tmpAppBkMk
	'bHasAccess = true
	sSql = "SELECT * FROM "& strTablePrefix & "APPS"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  do until rsA.eof
	    tmpAppID = tmpAppID & rsA("APP_ID") & "|"
	    tmpAppIName = tmpAppIName & rsA("APP_iNAME") & "|"
		tmpAppActive = tmpAppActive & rsA("APP_ACTIVE") & "|"
		tmpAppGroupsR = tmpAppGroupsR & rsA("APP_GROUPS_USERS") & "|"
		tmpAppGroupsW = tmpAppGroupsW & rsA("APP_GROUPS_WRITE") & "|"
		tmpAppGroupsF = tmpAppGroupsF & rsA("APP_GROUPS_FULL") & "|"
		tmpAppSubsc = tmpAppSubsc & rsA("APP_SUBSCRIPTIONS") & "|"
		tmpAppBkMk = tmpAppBkMk & rsA("APP_BOOKMARKS") & "|"
		tmpAppSecCode = tmpAppBkMk & rsA("APP_SUBSEC") & "|"
		rsA.movenext
	  loop
	  if tmpAppID <> "" then
		tmpAppID1 = split(tmpAppID,"|")
		tmpAppIName1 = split(tmpAppIName,"|")
		tmpAppActive1 = split(tmpAppActive,"|")
		tmpAppGroupsR1 = split(tmpAppGroupsR,"|")
		tmpAppGroupsW1 = split(tmpAppGroupsW,"|")
		tmpAppGroupsF1 = split(tmpAppGroupsF,"|")
		tmpAppSubsc1 = split(tmpAppSubsc,"|")
		tmpAppBkMk1 = split(tmpAppBkMk,"|")
		tmpAppSecCode1 = split(tmpAppSecCode,"|")
		acnt = ubound(tmpAppID1)-1
		redim arrAppPerms(acnt,8)
		for ag = 0 to acnt
		  arrAppPerms(ag,0) = tmpAppID1(ag)
		  arrAppPerms(ag,1) = tmpAppIName1(ag)
		  arrAppPerms(ag,2) = tmpAppActive1(ag)
		  arrAppPerms(ag,3) = tmpAppGroupsR1(ag)
		  arrAppPerms(ag,4) = tmpAppGroupsW1(ag)
		  arrAppPerms(ag,5) = tmpAppGroupsF1(ag)
		  arrAppPerms(ag,6) = tmpAppSubsc1(ag)
		  arrAppPerms(ag,7) = tmpAppBkMk1(ag)
		  arrAppPerms(ag,8) = tmpAppSecCode1(ag)
		next
	  end if
	else
	end if
end function



'##############################################
'##            NT Authentication             ##
'##############################################

sub NTUser()
	if Session(strCookieURL & "username")="" then
		strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strMemberTablePrefix & "MEMBERS.M_USERNAME, " & strMemberTablePrefix & "MEMBERS.M_NAME "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"

		set rs_chk = my_conn.Execute (strSql)

		if rs_chk.BOF or rs_chk.EOF then
			strLoginStatus = 0
		else
			Session(strCookieURL & "username") = rs_chk("M_NAME")
			DoCookies("true")
			strLoginStatus = 1
			strDBNTFUserName = rs_chk("M_NAME")
			strUserMemberID = clng(rs_chk("MEMBER_ID"))
			mLev = rs_chk("M_LEVEL")+1
			PMaccess = rs_chk("M_PMSTATUS")
			
			if rs_chk("M_PMRECEIVE") = 0 then
			  PMaccess = 0
			end if
			
			if mLev = 4 then 
				Session(strCookieURL & "Approval") = "256697926329"
			end if
		end if
		rs_chk.close	
		set rs_chk = nothing
	end if
end sub

function ChkAccountReg()
	' 
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_USERNAME "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'" 
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"

	set rs_chk = my_conn.Execute (strSql)

	if rs_chk.BOF or rs_chk.EOF then
		ChkAccountReg = "0"
	else
		ChkAccountReg = "1"
	end if

	rs_chk.close	
	set rs_chk = nothing

end function

sub NTAuthenticate()
	dim strUser, strNTUser, checkNT
	strNTUser = Request.ServerVariables("AUTH_USER") 
	strNTUser = replace(strNTUser, "\", "/")
	if Session(strCookieURL & "userid") = "" then
		strUser = Mid(strNTUser,(instr(1,strNTUser,"/")+1),len(strNTUser))
		Session(strCookieURL & "userid") = strUser
	end if
	if strNTGroups="1" then
		strNTGroupsSTR = Session(strCookieURL & "strNTGroupsSTR")
		if Session(strCookieURL & "strNTGroupsSTR") = "" then
			Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
			For Each strNTUserInfoGroup in strNTUserInfo.Groups
				strNTGroupsSTR=strNTGroupsSTR+", "+strNTUserInfoGroup.name
			NEXT
			Session(strCookieURL & "strNTGroupsSTR") = strNTGroupsSTR
		end if
	end if

	if strAutoLogon="1" then
		strNTUserFullName = Session(strCookieURL & "strNTUserFullName")
		if Session(strCookieURL & "strNTUserFullName") = "" then
			Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
			strNTUserFullName=strNTUserInfo.FullName
			Session(strCookieURL & "strNTUserFullName") = strNTUserFullName
		end if
	end if
end sub

'##############################################
'##         END - NT Authentication          ##
'##############################################
 %>
