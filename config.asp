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
<!-- #INCLUDE file="includes/classes/clsMenu.asp" -->
<!-- #include file="includes/classes/includes.asp" -->
<%
'on error resume next
Response.Buffer = true
dim startTime : startTime = timer
dim pageTimer, strDBType, strConnString, strTablePrefix, strMemberTablePrefix, strTheme, strWebMaster
dim intDisplay, FSOenabled, strShowImagePoweredBy
dim intAllowed, strCookieURL, strCurSymbol, showGames, showGold, showRep, sqlver
dim intMyMax, strUnicode, strUniqueID

'#################################################################################
'## SELECT YOUR DATABASE TYPE AND CONNECTION TYPE (access, sqlserver)
'#################################################################################
strDBType = "access"
'strDBType = "sqlserver"
'strDBType = "mysql"

	'## if you require unicode language support uncomment "YES" the line below
	'## and comment out the NO line. Unicode support is required for languages that use a different alphabet
	'## for more info see http://www.unicode.org/standard/WhatIsUnicode.html
	'## access is unicode by default and as such the variable will not be used
	strUnicode="NO"
	'strUnicode="YES"

'::: Provide the full path to your Access database here.
'strDBPath = "C:\Domains\your_folder\wwwroot\db\sp_db2k3.mdb"
strDBPath = server.MapPath("db/sp_db2k3.mdb")

'::: Choose one of the 3 connection strings below
'::: The string directly below is for Access DB. Do nothing if you are using Access
strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath '## MS Access 2000

'::: If you are using SQL Server, Comment out the line above and uncomment one
'::: of the 2 lines below and fill in the correct connection variables
 
'strConnString = "Provider=SQLOLEDB;Data Source=SQL_server_name_or_IP;Initial Catalog=db_name_here;UID=db_user_name_here;PWD=db_password_here" 'SQL Server

'strConnString = "Provider=SQLOLEDB;Data Source =SQL_server_name_or_IP;DSN=DSN_name;UID=SQL User;PWD=SQL password" 'SQL Server using DSN

'#################################################################################
':: strWebMaster is the list of Super Admins. Use lowercase member names.
':: names should always end with a comma
':: strWebMaster = "admin,skydogg,santa claus,"
strWebMaster = "administrator,"

':: Set the portal LCID 
':: the default of ENGLISH-USA is set
intPortalLCID = 1033

':: strUniqueID is your cookie name prefix. make it unique for your site. Keep it short.
':: If for some reason you need to "log everyone out", 
':: just change the variable below, save and upload.
strUniqueID = "SPRC2x"

':: installTheme is the default theme when you install the portal.
':: This is the folder name of the theme.
':: This variable also needs changed in site_setup.asp
installTheme = "Frosty_Sky"

':: If your site is in a VIRTUAL directory, 
':: or if you have problems logging in, uncomment
':: the line below and comment out the line under it.
'strCookieURL = "/"
strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))

':: Show page load time at bottom of all pages
pageTimer = 0	' 1 = yes; 0 = no   

':: Allow uploads?
':: This will override the database setting of "ON" for allowing uploads
intUploads = 1 ' 0 = OFF; 1 = ON

':: Allow subscriptions?
':: Having your EMAIL turned off will override this value
intSubscriptions = 1 ' 0=off 1=on

':: Allow bookmarks?
intBookmarks = 1 ' 0=off 1=on

':: registration form type
':: 0 = short form
':: 1 = long form
showRegisterLongForm = 0

':: Allow members to access the myMax feature:
':: If set to '0', only superadmin can arrange the front page
':: layout from using the myMax link in the members menu while
':: looking at the front page.
intMyMax = 0  '0=no 1=yes

'mLev for who can 'view source' on the editor
' 5 for super admin only
' 4 for all admins only
' 3 for moderators and admins
' 1 for all members
' 0 for members and guests
intEditor = 4 

':: HTML editor language.
':: use 2 letter abbreviation... lower case
strLang = "en"

'What HTML editor is used.
'SkyPortal currently supports FCKeditor and tinyMCE editor
editorType = "tinymce" 
'editorType = "fckeditor" 

'Currency Symbol
strCurSymbol = "$"

showGold = 0 	' 0=no; 1=yes;
showRep = 0 	' 0=no; 1=yes;
showGames = 0 	' 0=no; 1=yes;

'If installing on localhost, set the boolLocalHost value to true
'Uploads will not be available if this value is 'true'
boolLocalHost = false

':::::::: EVENTS CONFIG ::::::::::::::::::::::::::::::::::::::::::::::
' sets who is allowed to add new events to the calendar
intAllowed = 1	'All members
'intAllowed = 3	'Only Moderators and Admin
'intAllowed = 4	'Admins Only
intEventsUpcoming = 5  	'The number of Upcoming events to display
intEventsRecent = 5  	'The number of Recent events to display
intEventsToday = 5  	'The number of Today's events to display
':::::::: END EVENTS CONFIG ::::::::::::::::::::::::::::::::::::::::::

'#################################################################################
'## Do Not Edit Below This Line - It could destroy your database and lose data
'#################################################################################
strTablePrefix = "PORTAL_"
strMemberTablePrefix = "PORTAL_"

dim strSiteTitle, strCopyright, strTitleImage, strHomeURL, strWebSiteVersion
dim strAuthType, strForumStatus, strIPLogging
dim strEmail, strUniqueEmail, strMailMode, strMailServer, strSender
dim strDateType, strTimeAdjust, strTimeType, strForumTimeAdjust, strForumDateAdjust
dim strMoveTopicMode, strPrivateForums, strShowModerators, strShowRank, strAllowForumCode, strAllowHTML
dim strNoCookies, strEditedByDate
dim intHotTopicNum, strLockDown, strHotTopic
dim strIMGInPosts, strEmailVal
dim strHomepage, strICQ, strAIM, strSecureAdmin, strIcons, strGfxButtons
dim strBadWordFilter, strBadWords
dim strDefaultFontSize, strHeaderFontSize, strFooterFontSize
dim strRankColor1, strRankColor2, strRankColor3
dim strRankLevel0, strRankLevel1, strRankLevel2, strRankLevel3, strRankLevel4, strRankLevel5
dim intRankLevel0, intRankLevel1, intRankLevel2, intRankLevel3, intRankLevel4, intRankLevel5
dim strShowStatistics, strLogonForMail, strShowPaging, strPageSize, strPageNumberSize
dim strNTGroupsSTR, strPollCreate, strFeaturedPoll
dim strNewReg, pEnPrefix, blnSetup, my_Conn, strChkDate
dim strJokeOfTheWeek, strQuickReply, bFSOenabled, counter, strDBNTSQLName, okoame
dim strFloodCheck, strFloodCheckTime, strTimeLimit, strNavIcons, sysDebugMode
dim strMSN, strDefTheme, strAllowUploads,  strPMtype, strDBNTUserName
dim strICSLocation, strReminders, strIcalExist, strIcalNew, strForumSubscription
dim StrIPGateBan ,StrIPGateLck ,StrIPGateCok ,StrIPGateMet ,StrIPGateMsg ,StrIPGateLog ,StrIPGateTyp ,StrIPGateExp, StrIPGateCss, strIPGateVer, StrIPGateLkMsg, strIPGateNoAcMsg, StrIPGateWarnMsg, strHeaderType
dim strYAHOO, strFullName, strPicture, stMx, strSex, strCity, strState, strAge, strCountry, strOccupation
dim strBio, strHobbies, strLNews, strQuote, strMarStatus, strFavLinks, strRecentTopics, strAllowHideEmail
dim strUseExtendedProfile, strRankAdmin, strRankMod, strRankColorAdmin, strRankColorMod, strRankColor0
dim strRankColor4, strRankColor5, strNTGroups, strAutoLogon, strVar1, strVar2, strVar3, strVar4, strZip
dim strLoginType, browserReq, varBrowser, memID, isMAC, SecImage, dbHits, intMemberLCID, intPortalLCID
dim strCurDateAdjust, strCurDateString, strMTimeAdjust, strMTimeType, strMCurDateAdjust, strMCurDateString
dim mLev, strLoginStatus, strSiteOwner, strUserEmail, PMaccess, strUserMemberID
Dim arg1, arg2, arg3, arg4, arg5, arg6 'page breadcrumb variables
dim arrCurOnline(), arrGroups(), arrAppPerms()
		
strShowImagePoweredBy = "0"
intSkin=1
intIsSuperAdmin = 0
sysDebugMode = 0
pr=1
stMx = "Sk"
pEnPrefix = ""
strWebMaster = lcase(strWebMaster)

Session.LCID = intPortalLCID
'Session.LCID = 1033
	on error resume next
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear
	my_Conn.Open strConnString
	'	Lets check to see if the strConnString or db path has changed
		if my_conn.Errors.Count <> 0 then 
			'we can't connect, lets display the error
			my_conn.Errors.Clear 
			set my_Conn = nothing
			Response.Redirect "site_setup.asp?RC=1"
		end if
	on error goto 0

'blnSetup="N"
if request.QueryString("sky") = "dogg" then
  Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
end if

if blnSetup<>"Y" then
  if Application(strCookieURL & strUniqueID & "ConfigLoaded")= "" or IsNull(Application(strCookieURL & strUniqueID & "ConfigLoaded")) then 
	'## if the config variables aren't loaded into the Application object
	'## or after the admin has changed the configuration
	'## the variables get (re)loaded 
	
	'FileSystemObject check
	bFSOenabled = false
	if not boolLocalHost then
	 on error resume next
	 Err.Clear
	 set fso = Server.CreateObject("Scripting.FileSystemObject")
	 if err.number = 0 then
	   bFSOenabled = true
	   set fso = nothing
	 end if
	 Err.Clear
	 on error goto 0
	end if

	strSql = "SELECT C_STRSITETITLE "
	strSql = strSql & ", C_STRCOPYRIGHT "
	strSql = strSql & ", C_STRTITLEIMAGE "
	strSql = strSql & ", C_STRHOMEURL "
	strSql = strSql & ", C_STRAUTHTYPE "
	strSql = strSql & ", C_STREMAIL "
	strSql = strSql & ", C_STRUNIQUEEMAIL "
	strSql = strSql & ", C_STRMAILMODE "
	strSql = strSql & ", C_STRMAILSERVER "
	strSql = strSql & ", C_STRSENDER "
	strSql = strSql & ", C_STRDATETYPE "
	strSql = strSql & ", C_STRTIMEADJUST "
	strSql = strSql & ", C_STRTIMETYPE "
	strSql = strSql & ", C_STRMOVETOPICMODE "
	strSql = strSql & ", C_STRIPLOGGING "
	strSql = strSql & ", C_STRPRIVATEFORUMS "
	strSql = strSql & ", C_STRSHOWMODERATORS "
	strSql = strSql & ", C_STRALLOWFORUMCODE "
	strSql = strSql & ", C_STRALLOWHTML "
	strSql = strSql & ", C_STRNOCOOKIES "
	strSql = strSql & ", C_STRSECUREADMIN "
	strSql = strSql & ", C_STRHOTTOPIC "
	strSql = strSql & ", C_INTHOTTOPICNUM "
	strSql = strSql & ", C_STRIMGINPOSTS "
	strSql = strSql & ", C_STRHOMEPAGE "
	strSql = strSql & ", C_STRICQ "
	strSql = strSql & ", C_STRYAHOO "
	strSql = strSql & ", C_STRAIM "
	strSql = strSql & ", C_STRICONS "
	strSql = strSql & ", C_STRGFXBUTTONS "
	strSql = strSql & ", C_STREDITEDBYDATE "
	strSql = strSql & ", C_STRBADWORDFILTER "
	strSql = strSql & ", C_STRBADWORDS "
	strSql = strSql & ", C_STRSHOWRANK "
	strSql = strSql & ", C_STRRANKADMIN "
	strSql = strSql & ", C_STRRANKMOD "
	strSql = strSql & ", C_STRRANKLEVEL0 "
	strSql = strSql & ", C_STRRANKLEVEL1 "
	strSql = strSql & ", C_STRRANKLEVEL2 "
	strSql = strSql & ", C_STRRANKLEVEL3 "
	strSql = strSql & ", C_STRRANKLEVEL4 "
	strSql = strSql & ", C_STRRANKLEVEL5 "
	strSql = strSql & ", C_STRRANKCOLORADMIN "
	strSql = strSql & ", C_STRRANKCOLORMOD "
	strSql = strSql & ", C_STRRANKCOLOR0 "
	strSql = strSql & ", C_STRRANKCOLOR1 "
	strSql = strSql & ", C_STRRANKCOLOR2 "
	strSql = strSql & ", C_STRRANKCOLOR3 "
	strSql = strSql & ", C_STRRANKCOLOR4 "
	strSql = strSql & ", C_STRRANKCOLOR5 "
	strSql = strSql & ", C_INTRANKLEVEL0 "
	strSql = strSql & ", C_INTRANKLEVEL1 "
	strSql = strSql & ", C_INTRANKLEVEL2 "
	strSql = strSql & ", C_INTRANKLEVEL3 "
	strSql = strSql & ", C_INTRANKLEVEL4 "
	strSql = strSql & ", C_INTRANKLEVEL5 "
	strSql = strSql & ", C_STRSIGNATURES "
	strSql = strSql & ", C_STRSHOWSTATISTICS "
	strSql = strSql & ", C_STRLOGONFORMAIL "
	strSql = strSql & ", C_STRSHOWPAGING "
	strSql = strSql & ", C_STRPAGESIZE "
	strSql = strSql & ", C_STRPAGENUMBERSIZE "
	strSql = strSql & ", C_STRLOCKDOWN"
	strSql = strSql & ", C_STRFULLNAME"
	strSql = strSql & ", C_STRPICTURE"
	strSql = strSql & ", C_STRSEX"
	strSql = strSql & ", C_STRCITY"
	strSql = strSql & ", C_STRSTATE"
	strSql = strSql & ", C_STRZIP"
	strSql = strSql & ", C_STRAGE"
	strSql = strSql & ", C_STRCOUNTRY"
	strSql = strSql & ", C_STROCCUPATION"
	strSql = strSql & ", C_STRBIO"
	strSql = strSql & ", C_STRHOBBIES"
	strSql = strSql & ", C_STRLNEWS"
	strSql = strSql & ", C_STRQUOTE"
	strSql = strSql & ", C_STRMARSTATUS"
	strSql = strSql & ", C_STRFAVLINKS"
	strSql = strSql & ", C_STRRECENTTOPICS"
	strSql = strSql & ", C_STRHOMEPAGE"
	strSql = strSql & ", C_STRNTGROUPS"
	strSql = strSql & ", C_STRAUTOLOGON"
	strSql = strSql & ", C_STREMAILVAL"
	strSql = strSql & ", C_JOKEOFTHEWEEK"
	strSql = strSql & ", C_FORUMSTATUS"
	strSql = strSql & ", C_STRFLOODCHECK"
	strSql = strSql & ", C_STRFLOODCHECKTIME"
	strSql = strSql & ", C_POLLCREATE"
	strSql = strSql & ", C_FEATUREDPOLL"
	strSql = strSql & ", C_STRNEWREG"
	strSql = strSql & ", C_STRQUICKREPLY"
	strSql = strSql & ", C_STRMSN "
	strSql = strSql & ", C_STRDEFTHEME"
	strSql = strSql & ", C_ALLOWUPLOADS "
	strSql = strSql & ", C_PMTYPE"
	strSql = strSql & ", C_STRICSLOCATION"
	strSql = strSql & ", C_REMINDERS"
	strSql = strSql & ", C_ICALEXIST"
	strSql = strSql & ", C_ICALNEW"
	strSql = strSql & ", C_STRVAR1"
    strSql = strSql & ", C_STRVAR2"
    strSql = strSql & ", C_STRVAR3"
    strSql = strSql & ", C_STRVAR4"
    strSql = strSql & ", C_FORUMSUBSCRIPTION"
	' # added for IPGATE Mod
	strSql = strSql & ", C_STRIPGATEBAN"
	strSql = strSql & ", C_STRIPGATELCK"
	strSql = strSql & ", C_STRIPGATECOK"
	strSql = strSql & ", C_STRIPGATEMET"
	strSql = strSql & ", C_STRIPGATEMSG"
	strSql = strSql & ", C_STRIPGATELOG"
	strSql = strSql & ", C_STRIPGATETYP"
	strSql = strSql & ", C_STRIPGATEEXP"
	strSql = strSql & ", C_STRIPGATECSS"
	strSql = strSql & ", C_STRIPGATEVER"
	strSql = strSql & ", C_STRIPGATELKMSG"
	strSql = strSql & ", C_STRIPGATENOACMSG"
	strSql = strSql & ", C_STRIPGATEWARNMSG"
	strSql = strSql & ", C_STRHEADERTYPE"
	strSql = strSql & ", C_STRLOGINTYPE"
	strSql = strSql & ", C_SECIMAGE"
	strSql = strSql & ", C_INTSUBSKIN"
	strSql = strSql & ", C_ONEADAYDATE"
	strSql = strSql & " FROM " & strTablePrefix & "CONFIG "
	strSql = strSql & " WHERE CONFIG_ID = 1"
	
	
	on error resume next
	  set rsConfig = my_Conn.Execute(strSql)
	'	Lets check to see if the strConnString or db path has changed
		if err.number <> 0 then
			set my_Conn = nothing
			Response.Redirect "site_setup.asp?err=no_config_table"
		else
		
		end if
	on error goto 0

	Application.Lock
	application.Contents.RemoveAll()
	Application(strCookieURL & strUniqueID & "strSiteTitle") = rsConfig("C_STRSITETITLE")
	Application(strCookieURL & strUniqueID & "strCopyright") = rsConfig("C_STRCOPYRIGHT")
	Application(strCookieURL & strUniqueID & "strTitleImage") = rsConfig("C_STRTITLEIMAGE")
	Application(strCookieURL & strUniqueID & "strHomeURL") = rsConfig("C_STRHOMEURL")
	Application(strCookieURL & strUniqueID & "strAuthType") = rsConfig("C_STRAUTHTYPE")
	Application(strCookieURL & strUniqueID & "strEmail") = rsConfig("C_STREMAIL")
	Application(strCookieURL & strUniqueID & "strUniqueEmail") = rsConfig("C_STRUNIQUEEMAIL")
	Application(strCookieURL & strUniqueID & "strMailMode") = rsConfig("C_STRMAILMODE")
	Application(strCookieURL & strUniqueID & "strMailServer") = rsConfig("C_STRMAILSERVER")
	Application(strCookieURL & strUniqueID & "strSender") = rsConfig("C_STRSENDER")
	Application(strCookieURL & strUniqueID & "strDateType") = rsConfig("C_STRDATETYPE")
	Application(strCookieURL & strUniqueID & "strTimeAdjust") = rsConfig("C_STRTIMEADJUST")
	Application(strCookieURL & strUniqueID & "strTimeType") = rsConfig("C_STRTIMETYPE")
	Application(strCookieURL & strUniqueID & "strIPLogging") = rsConfig("C_STRIPLOGGING")
	Application(strCookieURL & strUniqueID & "strAllowForumCode") = rsConfig("C_STRALLOWFORUMCODE")
	Application(strCookieURL & strUniqueID & "strIMGInPosts") = rsConfig("C_STRIMGINPOSTS")
	Application(strCookieURL & strUniqueID & "strAllowHTML") = rsConfig("C_STRALLOWHTML")
	Application(strCookieURL & strUniqueID & "strNoCookies") = rsConfig("C_STRNOCOOKIES")
	Application(strCookieURL & strUniqueID & "strSecureAdmin") = rsConfig("C_STRSECUREADMIN")
	Application(strCookieURL & strUniqueID & "strLockDown") = rsConfig("C_STRLOCKDOWN")
	Application(strCookieURL & strUniqueID & "strIcons") = rsConfig("C_STRICONS")
	Application(strCookieURL & strUniqueID & "strGfxButtons") = rsConfig("C_STRGFXBUTTONS")
	Application(strCookieURL & strUniqueID & "strBadWordFilter") = rsConfig("C_STRBADWORDFILTER")
	Application(strCookieURL & strUniqueID & "strBadWords") = rsConfig("C_STRBADWORDS")
	Application(strCookieURL & strUniqueID & "strLogonForMail") = rsconfig("C_STRLOGONFORMAIL")
	Application(strCookieURL & strUniqueID & "STRNTGROUPS") = rsConfig("C_STRNTGROUPS")
	Application(strCookieURL & strUniqueID & "STRAUTOLOGON") = rsConfig("C_STRAUTOLOGON")
	Application(strCookieURL & strUniqueID & "strEmailVal") = rsConfig("C_STREMAILVAL")
	Application(strCookieURL & strUniqueID & "strFloodCheck") = rsConfig("C_STRFLOODCHECK")
	Application(strCookieURL & strUniqueID & "strFloodCheckTime") = rsConfig("C_STRFLOODCHECKTIME")
	Application(strCookieURL & strUniqueID & "strNewReg") = rsConfig("C_STRNEWREG")
	Application(strCookieURL & strUniqueID & "strDefTheme") = rsConfig("C_STRDEFTHEME")
	Application(strCookieURL & strUniqueID & "strAllowUploads") = rsConfig("C_ALLOWUPLOADS")
	Application(strCookieURL & strUniqueID & "strPMtype") = rsConfig("C_PMTYPE")
	Application(strCookieURL & strUniqueID & "STRIPGATEBAN")= rsConfig("C_STRIPGATEBAN")
	Application(strCookieURL & strUniqueID & "STRIPGATELCK")= rsConfig("C_STRIPGATELCK")
	Application(strCookieURL & strUniqueID & "STRIPGATECOK")= rsConfig("C_STRIPGATECOK")
	Application(strCookieURL & strUniqueID & "STRIPGATEMET")= rsConfig("C_STRIPGATEMET")
	Application(strCookieURL & strUniqueID & "STRIPGATEMSG")= rsConfig("C_STRIPGATEMSG")
	Application(strCookieURL & strUniqueID & "STRIPGATELOG")= rsConfig("C_STRIPGATELOG")
	Application(strCookieURL & strUniqueID & "STRIPGATETYP")= rsConfig("C_STRIPGATETYP")
	Application(strCookieURL & strUniqueID & "STRIPGATEEXP")= rsConfig("C_STRIPGATEEXP")
	Application(strCookieURL & strUniqueID & "STRIPGATECSS")= rsConfig("C_STRIPGATECSS")
	Application(strCookieURL & strUniqueID & "STRIPGATEVER")= rsConfig("C_STRIPGATEVER")
	Application(strCookieURL & strUniqueID & "STRIPGATELKMSG")= rsConfig("C_STRIPGATELKMSG")
	Application(strCookieURL & strUniqueID & "STRIPGATENOACMSG")= rsConfig("C_STRIPGATENOACMSG")
	Application(strCookieURL & strUniqueID & "STRIPGATEWARNMSG")= rsConfig("C_STRIPGATEWARNMSG")
	Application(strCookieURL & strUniqueID & "strHeaderType")= rsConfig("C_STRHEADERTYPE") 
	Application(strCookieURL & strUniqueID & "strLoginType")= rsConfig("C_STRLOGINTYPE")
	Application(strCookieURL & strUniqueID & "FSOenabled")= bFSOenabled 
	Application(strCookieURL & strUniqueID & "SECIMAGE")= rsConfig("C_SECIMAGE")
	Application(strCookieURL & strUniqueID & "intSubSkin")= rsConfig("C_INTSUBSKIN")
	Application(strCookieURL & strUniqueID & "strChkDate")= rsConfig("C_ONEADAYDATE")
	
	Application(strCookieURL & strUniqueID & "strMoveTopicMode") = rsConfig("C_STRMOVETOPICMODE")
	Application(strCookieURL & strUniqueID & "strPrivateForums") = rsConfig("C_STRPRIVATEFORUMS")
	Application(strCookieURL & strUniqueID & "strShowModerators") = rsConfig("C_STRSHOWMODERATORS")
	Application(strCookieURL & strUniqueID & "strHotTopic") = rsConfig("C_STRHOTTOPIC")
	Application(strCookieURL & strUniqueID & "intHotTopicNum") = rsConfig("C_INTHOTTOPICNUM")
	Application(strCookieURL & strUniqueID & "strShowRank") = rsConfig("C_STRSHOWRANK")
	Application(strCookieURL & strUniqueID & "strRankAdmin") = rsConfig("C_STRRANKADMIN")
	Application(strCookieURL & strUniqueID & "strRankMod") = rsConfig("C_STRRANKMOD")
	Application(strCookieURL & strUniqueID & "strRankLevel0") = rsConfig("C_STRRANKLEVEL0")
	Application(strCookieURL & strUniqueID & "strRankLevel1") = rsConfig("C_STRRANKLEVEL1")
	Application(strCookieURL & strUniqueID & "strRankLevel2") = rsConfig("C_STRRANKLEVEL2")
	Application(strCookieURL & strUniqueID & "strRankLevel3") = rsConfig("C_STRRANKLEVEL3")
	Application(strCookieURL & strUniqueID & "strRankLevel4") = rsConfig("C_STRRANKLEVEL4")
	Application(strCookieURL & strUniqueID & "strRankLevel5") = rsConfig("C_STRRANKLEVEL5")
	Application(strCookieURL & strUniqueID & "strRankColorAdmin") = rsConfig("C_STRRANKCOLORADMIN")
	Application(strCookieURL & strUniqueID & "strRankColorMod") = rsConfig("C_STRRANKCOLORMOD")
	Application(strCookieURL & strUniqueID & "strRankColor0") = rsConfig("C_STRRANKCOLOR0")
	Application(strCookieURL & strUniqueID & "strRankColor1") = rsConfig("C_STRRANKCOLOR1")
	Application(strCookieURL & strUniqueID & "strRankColor2") = rsConfig("C_STRRANKCOLOR2")
	Application(strCookieURL & strUniqueID & "strRankColor3") = rsConfig("C_STRRANKCOLOR3")
	Application(strCookieURL & strUniqueID & "strRankColor4") = rsConfig("C_STRRANKCOLOR4")
	Application(strCookieURL & strUniqueID & "strRankColor5") = rsConfig("C_STRRANKCOLOR5")
	Application(strCookieURL & strUniqueID & "intRankLevel0") = rsConfig("C_INTRANKLEVEL0")
	Application(strCookieURL & strUniqueID & "intRankLevel1") = rsConfig("C_INTRANKLEVEL1")
	Application(strCookieURL & strUniqueID & "intRankLevel2") = rsConfig("C_INTRANKLEVEL2")
	Application(strCookieURL & strUniqueID & "intRankLevel3") = rsConfig("C_INTRANKLEVEL3")
	Application(strCookieURL & strUniqueID & "intRankLevel4") = rsConfig("C_INTRANKLEVEL4")
	Application(strCookieURL & strUniqueID & "intRankLevel5") = rsConfig("C_INTRANKLEVEL5")
	Application(strCookieURL & strUniqueID & "strShowStatistics") = rsconfig("C_STRSHOWSTATISTICS")
	Application(strCookieURL & strUniqueID & "strShowPaging") = rsconfig("C_STRSHOWPAGING")
	Application(strCookieURL & strUniqueID & "strPageSize") = rsconfig("C_STRPAGESIZE")
	Application(strCookieURL & strUniqueID & "strPageNumberSize") = rsconfig("C_STRPAGENUMBERSIZE")
	Application(strCookieURL & strUniqueID & "strForumStatus") = rsConfig("C_FORUMSTATUS")
	Application(strCookieURL & strUniqueID & "strPollCreate") = rsConfig("C_POLLCREATE")
	Application(strCookieURL & strUniqueID & "strFeaturedPoll") = rsConfig("C_FEATUREDPOLL")
	Application(strCookieURL & strUniqueID & "strQuickReply") = rsConfig("C_STRQUICKREPLY")
	Application(strCookieURL & strUniqueID & "strForumSubscription") = rsConfig("C_FORUMSUBSCRIPTION")
	Application(strCookieURL & strUniqueID & "strEditedByDate") = rsConfig("C_STREDITEDBYDATE")
	Application(strCookieURL & strUniqueID & "strRecentTopics") = rsconfig("C_STRRECENTTOPICS")
	
	Application(strCookieURL & strUniqueID & "strICQ") = rsConfig("C_STRICQ")
	Application(strCookieURL & strUniqueID & "strYAHOO") = rsConfig("C_STRYAHOO")
	Application(strCookieURL & strUniqueID & "strAIM") = rsConfig("C_STRAIM")
	Application(strCookieURL & strUniqueID & "strMSN") = rsConfig("C_STRMSN")
	Application(strCookieURL & strUniqueID & "strHomepage") = rsConfig("C_STRHOMEPAGE")
	Application(strCookieURL & strUniqueID & "strFullName") = rsconfig("C_STRFULLNAME")
	Application(strCookieURL & strUniqueID & "strPicture") = rsconfig("C_STRPICTURE")
	Application(strCookieURL & strUniqueID & "strSex") = rsconfig("C_STRSEX")
	Application(strCookieURL & strUniqueID & "strAge") = rsconfig("C_STRAGE")
	Application(strCookieURL & strUniqueID & "strMarStatus") = rsconfig("C_STRMARSTATUS")
	Application(strCookieURL & strUniqueID & "strCity") = rsconfig("C_STRCITY")
	Application(strCookieURL & strUniqueID & "strState") = rsconfig("C_STRSTATE")
	Application(strCookieURL & strUniqueID & "strZip") = rsConfig("C_STRZIP")
	Application(strCookieURL & strUniqueID & "strCountry") = rsconfig("C_STRCOUNTRY")
	Application(strCookieURL & strUniqueID & "strOccupation") = rsconfig("C_STROCCUPATION")
	Application(strCookieURL & strUniqueID & "strFavLinks") = rsconfig("C_STRFAVLINKS")
	Application(strCookieURL & strUniqueID & "strVar1") = rsConfig("C_STRVAR1")
	Application(strCookieURL & strUniqueID & "strBio") = rsconfig("C_STRBIO")
	Application(strCookieURL & strUniqueID & "strVar2") = rsConfig("C_STRVAR2")
	Application(strCookieURL & strUniqueID & "strHobbies") = rsconfig("C_STRHOBBIES")
	Application(strCookieURL & strUniqueID & "strVar3") = rsConfig("C_STRVAR3")
	Application(strCookieURL & strUniqueID & "strLNews") = rsconfig("C_STRLNEWS")
	Application(strCookieURL & strUniqueID & "strVar4") = rsConfig("C_STRVAR4")
	Application(strCookieURL & strUniqueID & "strQuote") = rsconfig("C_STRQUOTE")
	
	Application(strCookieURL & strUniqueID & "strJokeOfTheWeek") = rsConfig("C_JOKEOFTHEWEEK")
	Application(strCookieURL & strUniqueID & "strICSLocation") = rsConfig("C_STRICSLOCATION")
	Application(strCookieURL & strUniqueID & "strReminders") = rsConfig("C_REMINDERS")
	Application(strCookieURL & strUniqueID & "strIcalExist") = rsConfig("C_ICALEXIST")
	Application(strCookieURL & strUniqueID & "strIcalNew") = rsConfig("C_ICALNEW")
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= "YES"

	Application.UnLock
  end if 
end if
okoame = 1
if blnSetup <> "Y" and stMx = "Sk" then 
	strSiteTitle = replace(Application(strCookieURL & strUniqueID & "strSiteTitle"),"''","'")
	strCopyright = Application(strCookieURL & strUniqueID & "strCopyright")
	strTitleImage = Application(strCookieURL & strUniqueID & "strTitleImage")
	strHomeURL = Application(strCookieURL & strUniqueID & "strHomeURL")
	strAuthType = Application(strCookieURL & strUniqueID & "strAuthType")
	strEmail = Application(strCookieURL & strUniqueID & "strEmail")
	strUniqueEmail = Application(strCookieURL & strUniqueID & "strUniqueEmail")
	strMailMode = Application(strCookieURL & strUniqueID & "strMailMode")
	strMailServer = Application(strCookieURL & strUniqueID & "strMailServer")
	strSender = Application(strCookieURL & strUniqueID & "strSender")
	strIPLogging = Application(strCookieURL & strUniqueID & "strIPLogging")
	strAllowForumCode = Application(strCookieURL & strUniqueID & "strAllowForumCode")
	strIMGInPosts = Application(strCookieURL & strUniqueID & "strIMGInPosts")
	strAllowHTML = Application(strCookieURL & strUniqueID & "strAllowHTML")
	strSecureAdmin = Application(strCookieURL & strUniqueID & "strSecureAdmin")
	strNoCookies = Application(strCookieURL & strUniqueID & "strNoCookies")
	strGfxButtons = Application(strCookieURL & strUniqueID & "strGfxButtons")
	strBadWordFilter = Application(strCookieURL & strUniqueID & "strBadWordFilter")
	strBadWords = Application(strCookieURL & strUniqueID & "strBadWords")
	strLockDown = Application(strCookieURL & strUniqueID & "strLockDown")
	strLogonForMail = Application(strCookieURL & strUniqueID & "strLogonForMail")
	
	strDateType = Application(strCookieURL & strUniqueID & "strDateType")
	strTimeAdjust = Application(strCookieURL & strUniqueID & "strTimeAdjust")
	strTimeType = Application(strCookieURL & strUniqueID & "strTimeType")
	
	strCurDateAdjust = DateAdd("h", strTimeAdjust , Now()) 'portal offset from server
	strCurDateString = DateToStr2(strCurDateAdjust)
	strForumTimeAdjust = strCurDateAdjust
	strForumDateAdjust = ChkDate2(strCurDateString)
	
	strMTimeAdjust = strTimeAdjust
	strMTimeType = strDateType
	strMCurDateAdjust = strCurDateAdjust
	strMCurDateString = strCurDateString
	intMemberLCID = intPortalLCID
	
	strNTGroups = Application(strCookieURL & strUniqueID & "STRNTGROUPS")
	strAutoLogon = Application(strCookieURL & strUniqueID & "STRAUTOLOGON")
	strEmailVal = Application(strCookieURL & strUniqueID & "STREMAILVAL")
	strFloodCheck = Application(strCookieURL & strUniqueID & "STRFLOODCHECK")
	strFloodCheckTime = Application(strCookieURL & strUniqueID & "STRFLOODCHECKTIME")
	strNewReg = Application(strCookieURL & strUniqueID & "STRNEWREG")
	strDefTheme = Application(strCookieURL & strUniqueID & "strDefTheme")
	strAllowUploads = Application(strCookieURL & strUniqueID & "strAllowUploads")
	strPMtype = Application(strCookieURL & strUniqueID & "strPMtype")
	StrIPGateBan = Application(strCookieURL & strUniqueID & "STRIPGATEBAN")
	StrIPGateLck = Application(strCookieURL & strUniqueID & "STRIPGATELCK")
	StrIPGateCok = Application(strCookieURL & strUniqueID & "STRIPGATECOK")
	StrIPGateMet = Application(strCookieURL & strUniqueID & "STRIPGATEMET")
	StrIPGateMsg = Application(strCookieURL & strUniqueID & "STRIPGATEMSG")
	StrIPGateLog = Application(strCookieURL & strUniqueID & "STRIPGATELOG")
	StrIPGateTyp = Application(strCookieURL & strUniqueID & "STRIPGATETYP")
	StrIPGateExp = Application(strCookieURL & strUniqueID & "STRIPGATEEXP")
	StrIPGateCss = Application(strCookieURL & strUniqueID & "STRIPGATECSS")
	strIPGateVer = Application(strCookieURL & strUniqueID & "STRIPGATEVER")
	StrIPGateLkMsg = Application(strCookieURL & strUniqueID & "STRIPGATELKMSG")
	strIPGateNoAcMsg = Application(strCookieURL & strUniqueID & "STRIPGATENOACMSG")
	StrIPGateWarnMsg = Application(strCookieURL & strUniqueID & "STRIPGATEWARNMSG") 
	strAllowHideEmail = "1"
	strUseExtendedProfile = true
	stWb = "yPor"
	strIcons = Application(strCookieURL & strUniqueID & "strIcons")
	strHeaderType = Application(strCookieURL & strUniqueID & "strHeaderType") 
	strLoginType = Application(strCookieURL & strUniqueID & "strLoginType") 
	FSOenabled = Application(strCookieURL & strUniqueID & "FSOenabled")
	SecImage = Application(strCookieURL & strUniqueID & "SecImage")
	intSubSkin = Application(strCookieURL & strUniqueID & "intSubSkin")
	strChkDate = Application(strCookieURL & strUniqueID & "strChkDate")
	'forums
	strMoveTopicMode = Application(strCookieURL & strUniqueID & "strMoveTopicMode")
	strPrivateForums = Application(strCookieURL & strUniqueID & "strPrivateForums")
	strShowModerators = Application(strCookieURL & strUniqueID & "strShowModerators")
	strHotTopic = Application(strCookieURL & strUniqueID & "strHotTopic")
	intHotTopicNum = Application(strCookieURL & strUniqueID & "intHotTopicNum")
	strEditedByDate = Application(strCookieURL & strUniqueID & "strEditedByDate")
	strShowRank = Application(strCookieURL & strUniqueID & "strShowRank")
	strRankAdmin = Application(strCookieURL & strUniqueID & "strRankAdmin")
	strRankMod = Application(strCookieURL & strUniqueID & "strRankMod")
	strRankLevel0 = Application(strCookieURL & strUniqueID & "strRankLevel0")
	strRankLevel1 = Application(strCookieURL & strUniqueID & "strRankLevel1")
	strRankLevel2 = Application(strCookieURL & strUniqueID & "strRankLevel2")
	strRankLevel3 = Application(strCookieURL & strUniqueID & "strRankLevel3")
	strRankLevel4 = Application(strCookieURL & strUniqueID & "strRankLevel4")
	strRankLevel5 = Application(strCookieURL & strUniqueID & "strRankLevel5")
	strRankColorAdmin = Application(strCookieURL & strUniqueID & "strRankColorAdmin")
	strRankColorMod = Application(strCookieURL & strUniqueID & "strRankColorMod")
	strRankColor0 = Application(strCookieURL & strUniqueID & "strRankColor0")
	strRankColor1 = Application(strCookieURL & strUniqueID & "strRankColor1")
	strRankColor2 = Application(strCookieURL & strUniqueID & "strRankColor2")
	strRankColor3 = Application(strCookieURL & strUniqueID & "strRankColor3")
	strRankColor4 = Application(strCookieURL & strUniqueID & "strRankColor4")
	strRankColor5 = Application(strCookieURL & strUniqueID & "strRankColor5")
	intRankLevel0 = Application(strCookieURL & strUniqueID & "intRankLevel0")
	intRankLevel1 = Application(strCookieURL & strUniqueID & "intRankLevel1")
	intRankLevel2 = Application(strCookieURL & strUniqueID & "intRankLevel2")
	intRankLevel3 = Application(strCookieURL & strUniqueID & "intRankLevel3")
	intRankLevel4 = Application(strCookieURL & strUniqueID & "intRankLevel4")
	intRankLevel5 = Application(strCookieURL & strUniqueID & "intRankLevel5")
	strShowStatistics = Application(strCookieURL & strUniqueID & "strShowStatistics")
	strShowPaging = Application(strCookieURL & strUniqueID & "strShowPaging")
	strPageSize = Application(strCookieURL & strUniqueID & "strPageSize")
	strPageNumberSize = Application(strCookieURL & strUniqueID & "strPageNumberSize")
	strForumStatus = Application(strCookieURL & strUniqueID & "strForumStatus") 
	strPollCreate = Application(strCookieURL & strUniqueID & "STRPOLLCREATE")
	strFeaturedPoll = Application(strCookieURL & strUniqueID & "STRFEATUREDPOLL")
	strQuickReply = Application(strCookieURL & strUniqueID & "STRQUICKREPLY")
	strForumSubscription = Application(strCookieURL & strUniqueID & "strForumSubscription") 
	'member stuff
	strFullName = Application(strCookieURL & strUniqueID & "strFullName")
	strPicture = Application(strCookieURL & strUniqueID & "strPicture")
	strMarStatus = Application(strCookieURL & strUniqueID & "strMarStatus")
	strAge = Application(strCookieURL & strUniqueID & "strAge")
	strSex = Application(strCookieURL & strUniqueID & "strSex")
	strCity= Application(strCookieURL & strUniqueID & "strCity")
	strState = Application(strCookieURL & strUniqueID & "strState")
	strZip = Application(strCookieURL & strUniqueID & "strZip")
	strCountry = Application(strCookieURL & strUniqueID & "strCountry") 
	strICQ = Application(strCookieURL & strUniqueID & "strICQ")
	strYAHOO = Application(strCookieURL & strUniqueID & "strYAHOO")
	strAIM = Application(strCookieURL & strUniqueID & "strAIM")
	strMSN = Application(strCookieURL & strUniqueID & "strMSN")
	strHomepage = Application(strCookieURL & strUniqueID & "strHomepage")
	strOccupation = Application(strCookieURL & strUniqueID & "strOccupation")
	strBio = Application(strCookieURL & strUniqueID & "strBio") 
	strHobbies = Application(strCookieURL & strUniqueID & "strHobbies") 
	strLNews = 	Application(strCookieURL & strUniqueID & "strLNews") 
	strQuote = Application(strCookieURL & strUniqueID & "strQuote") 
	strFavLinks = Application(strCookieURL & strUniqueID & "strFavLinks")
	strRecentTopics = Application(strCookieURL & strUniqueID & "strRecentTopics") 
	strVar1 = Application(strCookieURL & strUniqueID & "strVar1")
	strVar2 = Application(strCookieURL & strUniqueID & "strVar2")
	strVar3 = Application(strCookieURL & strUniqueID & "strVar3")
	strVar4 = Application(strCookieURL & strUniqueID & "strVar4")
	'events
	strICSLocation = Application(strCookieURL & strUniqueID & "strICSLocation")
	strReminders = Application(strCookieURL & strUniqueID & "strReminders")
	strIcalExist = Application(strCookieURL & strUniqueID & "strIcalExist")
	strIcalNew = Application(strCookieURL & strUniqueID & "strIcalNew")
	intSubSkin = Application(strCookieURL & strUniqueID & "intSubSkin")
	'intSubSkin = 0	

	if boolLocalHost then
		FSOenabled = false
		intUploads = 0
  		strAllowUploads = 0
	end if
	if not FSOenabled then
  		strAllowUploads = 0
	end if
	if strEmail = 0 then
	  intSubscriptions = 0
	end if

	if strSecureAdmin = "0" then
		strSecureAdmin = "1"
	end if

	'on error goto 0

	if strAuthType = "db" then
		strDBNTSQLName = "M_NAME"
		strAutoLogon ="0"
		strNTGroups  ="0"
	else
		strDBNTSQLName = "M_USERNAME"
	end if

'::::::::::::::::::::::: browser sniffer code. ::::::::::::::::::::::::::::::
	browserReq = request.ServerVariables("HTTP_USER_AGENT")
	varBrowser = ""
   		if instr(lcase(browserReq),"opera") <> 0 then
			varBrowser = "opera"
		elseif instr(lcase(browserReq),"firefox") <> 0 then ' Is FireFox browser
			varBrowser = "firefox"
		elseif instr(lcase(browserReq),"firebird") <> 0 then ' Is Firebird browser
			varBrowser = "firebird"
		elseif instr(lcase(browserReq),"safari") <> 0 then ' Is Safari browser
			varBrowser = "safari"
		elseif instr(lcase(browserReq),"lynx") <> 0 then ' Is lynx browser
			varBrowser = "lynx"
		elseif instr(lcase(browserReq),"camino") <> 0 then ' Is Safari browser
			varBrowser = "camino"
		elseif instr(lcase(browserReq),"msie") <> 0 then ' Is MSIE browser
			varBrowser = "ie"
		elseif instr(lcase(browserReq),"gecko") <> 0 then ' Is Netscape browser
			varBrowser = "netscape"
		else
			varBrowser = "other"	
		end if
			isMAC = false
			stPl = "tal."
		
		':: This code detects if the editor is browser compatable
		':: If not, then change settings to show the default [forum code] browser.
		'if (instr(lcase(browserReq),"mac") <> 0 and varBrowser <> "firefox") or varBrowser = "opera" then
		if instr(lcase(browserReq),"mac") <> 0 then
		  if varBrowser = "ie" then
			strAllowHtml = 0
			strAllowForumCode = 1
			strIMGInPosts = 1
		  end if
			isMAC = true
		end if

	':: Check the default theme to use on the site
	If trim(strDefTheme) = "" or isNull(strDefTheme) Then
  		strDefTheme = installTheme
	end if
	strTheme = strDefTheme
end if 
'response.Write("blnSetup: " & blnSetup)
%><!-- include file="lang/en/core.asp" -->
<% if blnSetup <> "Y" then %>
<%
if request("lang") <> "" then
  strLang = request("lang")
end if
if FSOenabled then
  include server.mappath("lang/" & strLang & "/core.asp")
else 
  %><!-- #include file="lang/en/core.asp" --><%
end if
 end if
%>