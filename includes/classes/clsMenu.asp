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

'******************************************************************************

'' @CLASSTITLE:		SkyMenu
'' @CREATOR:		Tom Nance (SkyDogg)		
'' @DESCRIPTION:	Custom menu class. Populates menu's from db. 
''					Vertical and horizontal menu's are supported.
''					Colors and graphics are controlled from the style_menus.css

'*******************************************************************************

Class SkyMenu

'********************************************************************************
	' Declare variables used in this class	'********************************************************************************
	Private p_sMenu					' [str] Name of menu to be pulled from the db
	Private p_sTitle				' [str] Themeblock Title
	Private p_iTemplate				' [int] Menu template to use
	Private p_shoExpanded
	Private p_canMinMax
	Private p_keepOpen
	Private icon_bar				' [str] counter
	Private p_thmBlk				' [int] Use themeblock  1=yes : 0=no
	Private p_createFile
	Private fly_menu								
	Private ie						' [int] counter
	Private ed						' [int] counter
	Private p_lnk					' [int] submenu id random counter
	Private mImg			
	Private cls
	Private icn
	Private alt		
	Private rCount					' [int] menu counter
	Private rsMnuTop
	Public mnuFile					' [str] path to menu file
	Public mnuTree					' [str] menu html
	Public mnuSubTree				' [str] flyout menu html
	Public mnuSubHTML				' [str] flyout menu html
	Private strTitles
	Private arHeadRow()
	'Private arSubTree()
	
'**********************************************************************************
	' Initialize default values for variables used in this class	'**********************************************************************************
	Private Sub Class_Initialize()
	  ie = 0
	  ed = 0
	  icon_bar = "<img src=""images/icons/icon_bar.gif"" align=""absmiddle"" height=""15"" width=""15"" border=""0"">&nbsp;"
	  mnuReset()
	End Sub
	
	Private Sub mnuReset()
	  cls = "none"
	  icn = "max1"
	  alt = txtExpand
		p_sMenu = ""
		p_sTitle = ""
		p_iTemplate = 4
		p_thmBlk = 0
		p_shoExpanded = 0
		p_canMinMax = 1
		p_keepOpen = 0
		p_lnk = 0
		mImg = ""
		fly_menu = ""
	    rCount = 0
		p_createFile = 0
		mnuFile = ""
		mnuTree = ""
		mnuSubTree = ""
		mnuSubHTML = ""
		strTitles = ":"
		dim arHeadRow(200,2)
	End Sub

'***********************************************************************************
	' Class Properties	'***********************************************************************************
	
	Public Property Get menuName()
		menuName = p_sMenu
	End Property
	Public Property Let menuName(v)
		p_sMenu = v
	End Property

	Public Property Get template()
		template = p_iTemplate
	End Property
	Public Property Let template(v)
		p_iTemplate = v
	End Property
	
	Public Property Get title()
		title = p_sTitle
	End Property
	Public Property Let title(v)
		p_sTitle = v
	End Property

	Public Property Get thmBlk()
		thmBlk = p_thmBlk
	End Property
	Public Property Let thmBlk(v)
		p_thmBlk = v
	End Property

	Public Property Get shoExpanded()
		shoExpanded = p_shoExpanded
	End Property
	Public Property Let shoExpanded(v)
		p_shoExpanded = v
	End Property

	Public Property Get canMinMax()
		canMinMax = p_canMinMax
	End Property
	Public Property Let canMinMax(v)
		p_canMinMax = v
	End Property

	Public Property Get keepOpen()
		keepOpen = p_keepOpen
	End Property
	Public Property Let keepOpen(v)
		p_keepOpen = v
	End Property 
	
	Public Property Let createFile(v)
	  if FSOenabled then
		p_createFile = v
	  end if
	End Property 
	

'***************************' Class Methods	'***************************
	Public sub GetMenu()
	  mnuStart()
	  
	  sSQL = "SELECT * from Menu Where Parent ='" & p_sMenu & "' and iName = '" & p_sMenu & "' order by mnuOrder asc"
	  Set rsMnuTop = Server.CreateObject("ADODB.Recordset")
  	  rsMnuTop.Open sSQL, my_Conn, 3, 1, &H0001
	  if not rsMnuTop.eof then
	 	if FSOenabled and p_createFile = 0 and FExists(mnuFile) then
		  'response.Write("From File")
		  include(mnuFile)
	   	  'mnuTree = ReadFile(mnuFile) 
	   	  'writeMenu(mnuTree)
	 	else
		  if FSOenabled then
		    p_createFile = 1
			if FExists(mnuFile) then
			  DelFile mnuFile
			end if
		  end if
		  
	      Select Case p_iTemplate
	   	    case 1
			  SimpleVMenu()
		    case 2
			  clickMenu()
		    case 3
			  fly_menu = "nav_menu"
			  hMenu()  'Horizontal navbar-type
		    case 4
			  clickMenu2(rsMnuTop)
		    case 5
			  fly_menu = "vfly_menu"
			  hMenu()
		    case 6
			  'Call ShowMenu6(p_sMenu)
		    case else
			  fly_menu = "nav_menu"
			  hMenu()
	      end select
		  
	 	end if
	  else
	    response.Write("<p>Menu not found</p>")
	  end if
	  
	  mnuEnd()
	end sub

'***************************' Class Methods	'***************************
	
	Private Function mnuStart()
	  ed = ed + 1
	  mnuFile = setMnuFile()
	  if p_thmBlk = 1 then
		if p_sTitle <> "" then
		  spThemeTitle = p_sTitle
		end if
		spThemeBlock1_open(intSkin)
	  end if
	End Function
	
	Private Function mnuEnd()
	  if p_thmBlk = 1 then
		spThemeBlock1_close(intSkin)
	  end if
	  mnuReset()
	End Function
	
	Private function getImageHTML(src,sw)
	   if src <> "" then
	     mImg = "<img src=""" & src & """ alt="""" border=""0"" hspace=""3"" />&nbsp;"
	   else
	     if sw = "" then
	       mImg = ""
		 else
	       mImg = sw
		 end if
	   end if
	end Function
	
	Public Function setMnuFile()
	  dim mf, tName
	  tName = p_sMenu &"_"& p_canMinMax &"_"& p_keepOpen &"_"& p_shoExpanded &"_"& p_iTemplate
	  mf = Server.MapPath("files/config/menu/" & tName & ".asp")
	  if p_createFile = 1 then
		DelFile mf
	  end if
	  setMnuFile = mf
	End Function
	
'************************* Simple Menu Template 1 ***************************	
	Public sub SimpleVMenu()
      mnuTree = mnuTree & "<div class=""menu"">" & vbCRLF
	  
	  sSQL = "SELECT * from Menu Where Parent ='" & p_sMenu & "' and iName = '" & p_sMenu & "' order by mnuOrder asc"
	  set rsParent = my_Conn.Execute(sSQL)
	  do while not rsParent.eof
	    if trim(rsParent("Link")) <> "" then
		  mnuTree = mnuTree & "<a href=""" & rsParent("Link") & """ target=""" & rsParent("Target") & """>"
		elseif trim(rsParent("onClick")) <> "" then
		  mnuTree = mnuTree & "<a href=""javascript:;"" onClick=""" & trim(rsParent("onClick")) & """>"
		else
		  mnuTree = mnuTree & "<a href=""javascript:;"">"
		end if
		  mnuTree = mnuTree & "&nbsp;-&nbsp;" & rsParent("Name") & "<br /></a>"
		rsParent.movenext
	  loop
	  set rsParent = nothing
	  
      mnuTree = mnuTree & "</div>" & vbCRLF
	  writeMenu(mnuTree)
	end sub
	
'***************************** Menu Template 1 ********************************	
	Public sub clickMenu2(oRs)
	  mnuSubHTML = ""
	  if p_shoExpanded = 1 then
	    cls = "block"
	    icn = "min1"
	    alt = txtCollapse
	  end if
   	  Randomize()
  	  ed=Int(Rnd()*9000)
	  mnuTree = "<div id=""masterdiv" & ed & """ class=""mnuContainer"" style=""text-align:left;"">" & vbCRLF
	  do while not oRs.eof   ' AND NAME = '" & oRs("NAME") & "'
	    if oRs("mnuAdd") <> "" then
		  addMenuGroup(oRs("mnuAdd"))
		else
	   	  writeMenuGroup(oRs)
		end if
	   rsMnuTop.MoveNext 
	  loop 
      mnuTree = mnuTree & "</div>" & vbCRLF 
	  mnuTree = mnuTree & mnuSubHTML
	  writeMenu(mnuTree)	
	end sub
	
	Private Function addMenuGroup(mAdd)
	  	  sSQL = "SELECT * from Menu Where Parent ='" & mAdd & "' and iName = '" & mAdd & "' order by mnuOrder asc"
	  	  Set rsAdd = Server.CreateObject("ADODB.Recordset")
  	  	  rsAdd.Open sSQL, my_Conn, 3, 1, &H0001
	  	  do while not rsAdd.eof
		    
		    if rsAdd("mnuAdd") <> "" then
	  	  	  sSQL = "SELECT * from Menu Where Parent ='" & rsAdd("mnuAdd") & "' and iName = '" & rsAdd("mnuAdd") & "' order by mnuOrder asc"
	  	  	  Set rsAd = Server.CreateObject("ADODB.Recordset")
  	  	  	  rsAd.Open sSQL, my_Conn, 3, 1, &H0001
	  	  	  do while not rsAd.eof
		    
		    	if rsAd("mnuAdd") <> "" then
	  	  	  	  sSQL = "SELECT * from Menu Where Parent ='" & rsAd("mnuAdd") & "' and iName = '" & rsAd("mnuAdd") & "' order by mnuOrder asc"
	  	  	  	  Set rsA = Server.CreateObject("ADODB.Recordset")
  	  	  	  	  rsA.Open sSQL, my_Conn, 3, 1, &H0001
	  	  	  	  do while not rsA.eof
			      	writeMenuGroup(rsA)
				  	rsA.movenext
		  	  	  loop
		  	  	  set rsA = nothing
				else
			  	  writeMenuGroup(rsAd)
				end if
			
				rsAd.movenext
		  	  loop
		  	  set rsAd = nothing
			else
			  writeMenuGroup(rsAdd)
			end if
			
			rsAdd.movenext
		  loop
		  set rsAdd = nothing
	End Function
	
	Private Function accessStart(a)
	  aStart = ""
	  if a <> "" then
	      aStart = "<% if hasAccess(""" & a & """) then %" & ">"
	  end if
	  accessStart = aStart
	End Function
	
	Private Function accessEnd(a)
	  aEnd = ""
	  if a <> "" then
	    aEnd = "<% end if %" & ">"
	  end if
	  accessEnd = aEnd
	End Function
	
	Private Function appStart(a)
	  aStart = ""
	  if a <> "" then
	   if cint(a) > 0 then
	    'sSql = "select APP_iNAME from " & strTablePrefix & "_APPS where APP_ID = " & a
	    sSql = "select APP_iNAME from PORTAL_APPS where APP_ID = " & a
		set rsA = my_Conn.execute(sSql)
		if not rsA.eof then
	      aStart = "<% if chkApp(""" & rsA("APP_iNAME") & """,""USERS"") then %" & ">"
		end if
		set rsA = nothing
	   end if
	  end if
	  appStart = aStart
	End Function
	
	Private Function appEnd(ax)
	  aEnd = ""
	  if ax <> "" then
	   if ax > 0 then
	    aEnd = "<% end if %" & ">"
	   end if
	  end if
	  appEnd = aEnd
	End Function
		 
	Private Function setMnuFunction(f)
	  if trim(f) <> "" then
	    if p_createFile = 1 then
		else
		  'execute("Call " & f)
		end if
		setMnuFunction = "<%= " & f & " %" & ">"
	  end if
	End Function
	
	sub writeMenuGroup(oRs)
	    ie = ie + 1
   		Randomize()
  		p_lnk=Int(Rnd()*9000)
	    strSQL = "SELECT COUNT(*) FROM Menu Where Parent ='" & oRs("Name") & "' and iName = '" & oRs("iName") & "'"
		set rsCount = my_Conn.Execute(strSQL)
		intCount = clng(rsCount(0))
		set rsCount = nothing 
		
	    getImageHTML oRs("mnuImage"),"&nbsp;&nbsp;"
		
	    mnuTree = mnuTree & appStart(oRs("app_id"))
	    mnuTree = mnuTree & accessStart(oRs("mnuAccess"))
	   If oRs("Link") <> "" or oRs("onClick") <> "" Then
		 'mnuTree = mnuTree & "<a href=""" & rsParent("Link") & """ target=""" & rsParent("Target") & """>"
	     mnuTree = mnuTree & "<div class=""mnuHead"" "
		 mnuTree = mnuTree & "onMouseOver=""this.className='mnuHeadHover';"" onMouseOut=""this.className='mnuHead';"" style=""cursor:pointer;"""
		 If trim(oRs("Link")) <> "" then
		  if oRs("Target") = "_parent" then
		    mnuTree = mnuTree & " onClick=""javascript:window.location = '" & oRs("Link") & "';"""
		  else
		    mnuTree = mnuTree & " onClick=""window.open('" & oRs("Link") & "')"""
		  end if
		 elseif trim(oRs("onClick")) <> "" Then
		    mnuTree = mnuTree & " onClick=""" & replace(oRs("onClick"),"''","'") & """"
		 end if
		 mnuTree = mnuTree & ">" & mImg
		 mnuTree = mnuTree & "<b>" & oRs("Name") & "</b>"
		 
		 mnuTree = mnuTree & setMnuFunction(oRs("mnuFunction"))
			   'if oRs("mnuFunction") <> "" then
			   '  execute("Call " & oRs("mnuFunction"))
			   'end if
		 mnuTree = mnuTree & "</div>"
	   Else
	   	 if intCount = 0 then
	       mnuTree = mnuTree & "<div class=""mnuHead"" onMouseOver=""this.className='mnuHeadHover';"" onMouseOut=""this.className='mnuHead';"" style=""cursor:pointer;"">"
      	   mnuTree = mnuTree & mImg & "<b>" & oRs("Name") & "</b>"
		   mnuTree = mnuTree & setMnuFunction(oRs("mnuFunction"))
			   'if oRs("mnuFunction") <> "" then
			   '  execute("Call " & oRs("mnuFunction"))
			   'end if
		   mnuTree = mnuTree & "</div>" & vbCRLF
		 Else
	       icn = "max1"
	       mnuTree = mnuTree & "<div class=""mnuHead"""
		   if p_canMinMax = 1 then
		     mnuTree = mnuTree & " style=""cursor:pointer;"" onMouseOver=""this.className='mnuHeadHover';"" onMouseOut=""this.className='mnuHead';"""
			 if p_keepOpen = 1 then
			   mnuTree = mnuTree & " onclick=""javascript:mwpHSa('block" & p_lnk & "','2');"""
			 else
			   mnuTree = mnuTree & " onclick=""SwitchMenu('masterdiv" & ed & "','block" & p_lnk & "')"""
			 end if
		   end if
		   mnuTree = mnuTree & ">"
		   if p_canMinMax = 1 then
		     mnuTree = mnuTree & "<img name=""block" & p_lnk & "Img"" id=""block" & p_lnk & "Img"" src=""Themes/" & chr(60) & chr(37) & chr(61) & " strtheme " & chr(37) & chr(62) & "/icon_" & icn & ".gif"" vspace=""2"" align=""right"" style=""cursor:pointer;"" title=""" & alt & """ alt=""" & alt & """>"
		   end if
      	   mnuTree = mnuTree & mImg & "<b>" & oRs("Name") & "</b>"
		   
		   mnuTree = mnuTree & setMnuFunction(oRs("mnuFunction"))
			   'if oRs("mnuFunction") <> "" then
			   '  execute("Call " & oRs("mnuFunction"))
			   'end if
		   mnuTree = mnuTree & "</div>"
		   	   
		   set rsChild = my_Conn.Execute("SELECT * from Menu Where Parent ='" & oRs("Name") & "' and iName = '" & oRs("iName") & "' order by mnuOrder asc")
		   if not rsChild.eof then
      		 'mnuTree = mnuTree & "<div class=""menuX"" id=""block" & p_lnk & """ style=""display:" & cls & ";"">" & vbCRLF
      		 mnuTree = mnuTree & "<span class=""submenu"" id=""block" & p_lnk & """ style=""display:" & cls & ";"">" & vbCRLF
             do while not rsChild.eof
			   mnuTree = mnuTree & accessStart(rsChild("mnuAccess"))
			   If rsChild("Link") <> "" or rsChild("onClick") <> "" Then
		 	   	   mnuTree = mnuTree & "<div class=""mnuChild"" onMouseOver=""this.className='mnuChildHover';"" onMouseOut=""this.className='mnuChild';"""
			       mnuTree = mnuTree & " style=""cursor:pointer;"""
			     If rsChild("Link") <> "" then
			       if rsChild("Target") = "_parent" then
				   mnuTree = mnuTree & " onClick=""javascript:window.location = '" & rsChild("Link") & "';"""
				   else
			       mnuTree = mnuTree & " onClick=""window.open('" & rsChild("Link") & "')"""
				   end if
				 elseif rsChild("onClick") <> "" Then
		    	   mnuTree = mnuTree & " onClick=""" & replace(rsChild("onClick"),"''","'") & """"
				 end if
			   else
			     clickFlyRow(rsChild)
		 	   	   'mnuTree = mnuTree & "<div class=""mnuChild"" onMouseOver=""this.className='mnuChildHover';"" onMouseOut=""this.className='mnuChild';"""
			       'mnuTree = mnuTree & " style=""cursor:pointer;"""
			   end if
			   mnuTree = mnuTree & ">" & vbCRLF
			   getImageHTML rsChild("mnuImage"),icon_bar
		  	   mnuTree = mnuTree & mImg
		  	   mnuTree = mnuTree & rsChild("Name")
		       mnuTree = mnuTree & setMnuFunction(rsChild("mnuFunction"))
			   mnuTree = mnuTree & "</div>" & vbCRLF
	   		   mnuTree = mnuTree & accessEnd(rsChild("mnuAccess"))
               rsChild.MoveNext 
		     loop
			 mnuTree = mnuTree & "</span>" & vbCRLF
			 'mnuTree = mnuTree & "</div>" & vbCRLF
           End if 
         end if
	     set rsChild = nothing
       End If
	   mnuTree = mnuTree & accessEnd(oRs("mnuAccess"))
	   mnuTree = mnuTree & appEnd(oRs("app_id"))
	end sub
	
'***************************** Menu Template 3 ********************************	
	Public sub clickMenu()
	  mnuTree = "<div id=""masterdiv" & ed & """>" & vbCRLF
      mnuTree = mnuTree & "<div class=""menu"">" & vbCRLF
	  strSQL = "SELECT * from Menu Where Parent ='" & p_sMenu & "' and iName = '" & p_sMenu & "' order by mnuOrder asc"
	  set rsParent = my_Conn.Execute(strSQL)
	  do while not rsParent.eof 
	    ie = ie + 1
	    strSQL = "SELECT COUNT(*) FROM Menu Where Parent ='" & rsParent("Name") & "' and iName = '" & p_sMenu & "'"
		set rsCount = my_Conn.Execute(strSQL)
		intCount = clng(rsCount(0))
		set rsCount = nothing 
	  
	   If rsParent("Link") <> "" Then
        'mnuTree = mnuTree & "<div class=""menu"">"
		mnuTree = mnuTree & "<a href=""" & rsParent("Link") & """ target=""" & rsParent("Target") & """>" & rsParent("Name") & "<br /></a>"
	    'mnuTree = mnuTree & "</div>" & vbCRLF
	   Else
	   	 if intCount = 0 then
           'mnuTree = mnuTree & "<div class=""menu"">" & vbCRLF
		   mnuTree = mnuTree & "<a href=""javascript:;"">" & rsParent("Name") & "<br /></a>" & vbCRLF  
	       'mnuTree = mnuTree & "</div>" & vbCRLF
		 Else
           'mnuTree = mnuTree & "<div class=""menu"">" & vbCRLF
		   mnuTree = mnuTree & "<a href=""javascript:;"" onclick=""SwitchMenu('masterdiv" & ed & "','subH" & ie & "')"">" & rsParent("Name") & "<br /></a>" & vbCRLF
	       'mnuTree = mnuTree & "</div>" & vbCRLF
		   	   
		   set rsChild = my_Conn.Execute("SELECT * from Menu Where Parent ='" & rsParent("Name") & "' and iName = '" & p_sMenu & "' order by mnuOrder asc")
		   if not rsChild.eof then
             mnuTree = mnuTree & "<span class=""submenu"" id=""subH" & ie & """>" & vbCRLF
             do while not rsChild.eof
		 	   'mnuTree = mnuTree & "<div class=""menu"">" & vbCRLF
			   if rsChild("mnuImage") <> "" then
		  		 mnuTree = mnuTree & "<a href=""" & rsChild("Link") & """ target=""" & rsChild("Target") & """>"
		  		 mnuTree = mnuTree & "<img src=""images/" & rsChild("mnuImage") & """ height=""12"" width=""12"" alt="""" border=""0"" />"
		  		 mnuTree = mnuTree & "&nbsp;" & rsChild("Name") & "</a>"
			   Else
		 		 mnuTree = mnuTree & "<a href=""" & rsChild("Link") & """ target=""" & rsChild("Target") & """>&nbsp;&nbsp;- " & rsChild("Name") & "</a>"
			   End If
			   'mnuTree = mnuTree & "<div>" & vbCRLF
               rsChild.MoveNext 
		     loop
			 mnuTree = mnuTree & "</span>" & vbCRLF
           End if 
         end if
	     set rsChild = nothing
       End If
	   rsParent.MoveNext 
	  loop 
	  set rsParent = nothing
      mnuTree = mnuTree & "</div></div>" & vbCRLF
	  writeMenu(mnuTree)
	end sub
'***************************** Menu Template 4 ****************************	
	Public sub ShowMenu4(title)
	
	end sub
'**************************** Menu Template 5 *****************************	
	Public sub ShowMenu5(title)

	end sub
'**************************** Menu Template 6 ********************************
	
	Public sub clickFlyRow(oRs)
	    'ed = ed + 1
		redim arHeadRow(200,2)
	  'for mx = 0 to (rCount - 1)
	  mx = 0
	  fly_menu = "vfly_menu"
		  mnuName = oRs("Name")
		  mnuLink = oRs("Link")
		  ConClick = oRs("onClick")
		  CaddMenu = oRs("mnuAdd")
		  Ciname = oRs("iNAME")
		  Cimg = oRs("mnuImage")
		  Cfunct = oRs("mnuFunction")
			'response.Write("CaddMenu1: " & CaddMenu & "<br>")
			'response.Write("Ciname1: " & Ciname & "<br><br>")
		if CaddMenu <> "" then
	  	  sSQL = "SELECT * from Menu Where Parent ='" & CaddMenu & "' and iName = '" & CaddMenu & "' order by mnuOrder asc"
	  	  Set rsSb = my_Conn.execute(sSql)
	  	  if not rsSb.eof then
		    do until rsSb.eof
			  'redim preserve arHeadRow(mx + 1,2)
		  	  mnuName = rsSb("Name")
		  	  mnuLink = rsSb("Link")
		  	  ConClick = rsSb("onClick")
		  	  CaddMenu = rsSb("mnuAdd")
		  	  Ciname = rsSb("iNAME")
		  	  Cimg = rsSb("mnuImage")
		  	  Cfunct = rsSb("mnuFunction")
			'response.Write("CaddMenu2: " & CaddMenu & "<br>")
			'response.Write("Ciname2: " & Ciname & "<br><br>")
		      if CaddMenu <> "" then
			  
	  	  		sSQL = "SELECT * from Menu Where Parent ='" & CaddMenu & "' and iName = '" & CaddMenu & "' order by mnuOrder asc"
	  	  		Set rsS = my_Conn.execute(sSql)
	  	  		if not rsS.eof then
		    	  do until rsS.eof
			  		'redim preserve arHeadRow(mx + 1,2)
		  	  		mnuName = rsS("Name")
		  	  		mnuLink = rsS("Link")
		  	  		ConClick = rsS("onClick")
		  	  		CaddMenu = rsS("mnuAdd")
		  	  		Ciname = rsS("iNAME")
		  	  		Cparent = rsS("Parent")
		  	  		Cimg = rsS("mnuImage")
		  	  		Cfunct = rsS("mnuFunction")
			'response.Write("CaddMenu3: " & CaddMenu & "<br>")
			'response.Write("Ciname3: " & Ciname & "<br><br>")
		      		if CaddMenu <> "" then
			    	  writeSubHeaderV CaddMenu,Cfunct,Cimg,CaddMenu
		  			  mx = mx + 1
			  		else
			    	  writeSubHeaderV mnuName,Cfunct,Cimg,Ciname
		  			  mx = mx + 1
			  		end if
			  		rsS.movenext
				  loop
		  		end if
		  		set rsS = nothing
		  
			  else
			    writeSubHeaderV mnuName,Cfunct,Cimg,Ciname
		  		mx = mx + 1
			  end if
			  rsSb.movenext
			loop
		  end if
		  set rsSb = nothing
		else
		  writeSubHeaderV mnuName,Cfunct,Cimg,Ciname
		  mx = mx + 1
		  
		end if
	  'next
		'mnuTree = mnuTree & "></div"
		
		
	  for sm = 0 to ubound(arHeadRow)
		if trim(arHeadRow(sm,0)) <> "" then
		  C_MenuTree arHeadRow(sm,0),arHeadRow(sm,1),arHeadRow(sm,2)
		end if
	  next
	  
	  'mnuTree = mnuTree & mnuSubHTML
	  'writeMenu(mnuTree)
	  'writeMenu(mnuSubHTML)
	  
	end sub
	
	Private Function writeSubHeaderV(mName,mFunct,img,mAM)
   		    Randomize()
  		    p_lnk=Int(Rnd()*9000)
		 	   mnuTree = mnuTree & "<div class=""mnuChild"" onMouseOver=""buttonMouseover2(event, 'sub_" & p_lnk & "','vfly_menu');"" style=""cursor:pointer"">"
			   mnuTree = mnuTree & "<img src=""images/tri.gif"" align=""right"" class=""menuItemArrow"" border=""0"" vspace=""3"" hspace=""5"""
			
			arHeadRow(0,0) = mName
			arHeadRow(0,1) = "sub_" & p_lnk
			arHeadRow(0,2) = mAM
	End Function
		
	Private function C_MenuTree(parent,mnID,iNam)
	  mnuSubTree = ""
	  sSQL = "SELECT * from Menu Where Parent ='" & parent & "' and iName = '" & iNam & "' order by mnuOrder asc"
	  Set rsSub = Server.CreateObject("ADODB.Recordset")
  	  rsSub.Open sSQL, my_Conn, 3, 1, &H0001
	  if not rsSub.eof then
  	    rSCount = rsSub.recordcount
		'redim arSubTree(rSCount,2)
        mnuSubHTML = mnuSubHTML & "<div id=""" & mnID & """ class=""" & fly_menu & """ onmouseover=""menuMouseover(event)"">" & vbCRLF
	    'do while not rsSub.eof 
		for mx = 0 to (rSCount - 1)
		  mnuSubHTML = mnuSubHTML & accessStart(rsSub("mnuAccess"))
		  mnuName = rsSub("Name")
		  mnuLink = rsSub("Link")
		  ConClick = rsSub("onClick")
		  m_funct = rsSub("mnuFunction")
		  m_targ = rsSub("Target")
		  'mnuSubHTML = mnuSubHTML & setMnuFunction(m_funct)
			if mnuLink <> "" then
			    mnuSubHTML = mnuSubHTML & "<a href=""" & mnuLink & """ class=""" & fly_menu & "Item"" target=""" & m_targ & """>" & mnuName & "<br></a>" & vbCRLF
			elseif ConClick <> "" then
			    mnuSubHTML = mnuSubHTML & "<a href=""javascript:;"" onClick=""" & replace(ConClick,"''","'") & """ class=""" & fly_menu & "Item"">" & mnuName & "<br></a>" & vbCRLF
			else
   		      Randomize()
  		      p_lnk=Int(Rnd()*9000)
			  mnuSubHTML = mnuSubHTML & "<a class=""" & fly_menu & "Item"" href=""javascript:;"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'sub_" & p_lnk & "');""><span class=""menuItemText"">" & mnuName & "</span>"
			  mnuSubHTML = mnuSubHTML & "<span class=""menuItemArrow"">&#9654;</span><br></a>" & vbCRLF
			  mnuSubTree = mnuSubTree & mnuName & ":sub_" & p_lnk & ":" & rsSub("iName") & "|"
			end if
	   	  mnuSubHTML = mnuSubHTML & accessEnd(rsSub("mnuAccess"))
		  rsSub.movenext
		next
		mnuSubHTML = mnuSubHTML & setMnuFunction(m_funct)
        mnuSubHTML = mnuSubHTML & "</div>" & vbCRLF & vbCRLF
		'writeMenu(mnuTree)
		
		if mnuSubTree <> "" then
		  if instr(mnuSubTree,"|") <> 0 then
		    aTree = split(mnuSubTree,"|")
			for xm = 0 to ubound(aTree)-1
			  'response.Write(aTree(xm))
			  bTree = split(aTree(xm),":")
			  C_MenuTree bTree(0),bTree(1),bTree(2)
			next
		  else
		  end if
		end if
	  else
	    'response.Write("Menu not found")
	  end if
	  set rsSub = nothing
	end function
	
	Public sub hMenu()
	  'mnuReset()
  	    rCount = rsMnuTop.recordcount
	    'ed = ed + 1
		'redim arHeadRow(rCount,1)
		redim arHeadRow(200,2)
		if p_iTemplate = 5 then
		  mnuTree = "<div class=""menuBarV"">"
		else
		  mnuTree = "<div class=""menuBar"" style=""width:100%;"">" & vbCRLF
		end if
	  'for mx = 0 to (rCount - 1)
	  mx = 0
	  do until rsMnuTop.eof
		  mnuName = rsMnuTop("Name")
		  mnuLink = rsMnuTop("Link")
		  ConClick = rsMnuTop("onClick")
		  CaddMenu = rsMnuTop("mnuAdd")
		  Ciname = rsMnuTop("iNAME")
		  CappID = rsMnuTop("app_id")
		  Caccess = rsMnuTop("mnuAccess")
		  Cfunct = rsMnuTop("mnuFunction")
		  
	      mnuTree = mnuTree & appStart(CappID)
		
			'response.Write("CaddMenu1: " & CaddMenu & "<br>")
			'response.Write("Ciname1: " & Ciname & "<br><br>")
		if CaddMenu <> "" then
	  	  sSQL = "SELECT * from Menu Where Parent ='" & CaddMenu & "' and iName = '" & CaddMenu & "' order by mnuOrder asc"
	  	  Set rsSb = my_Conn.execute(sSql)
	  	  if not rsSb.eof then
		    do until rsSb.eof
			  'redim preserve arHeadRow(mx + 1,2)
		  	  mnuName = rsSb("Name")
		  	  mnuLink = rsSb("Link")
		  	  ConClick = rsSb("onClick")
		  	  CaddMenu = rsSb("mnuAdd")
		  	  Ciname = rsSb("iNAME")
			  Caccess = rsSb("mnuAccess")
			  Cfunct = rsSb("mnuFunction")
			'response.Write("CaddMenu2: " & CaddMenu & "<br>")
			'response.Write("Ciname2: " & Ciname & "<br><br>")
		      if CaddMenu <> "" then
			  
	  	  		sSQL = "SELECT * from Menu Where Parent ='" & CaddMenu & "' and iName = '" & CaddMenu & "' order by mnuOrder asc"
	  	  		Set rsS = my_Conn.execute(sSql)
	  	  		if not rsS.eof then
		    	  do until rsS.eof
			  		'redim preserve arHeadRow(mx + 1,2)
		  	  		mnuName = rsS("Name")
		  	  		mnuLink = rsS("Link")
		  	  		ConClick = rsS("onClick")
		  	  		CaddMenu = rsS("mnuAdd")
		  	  		Ciname = rsS("iNAME")
		  	  		Cparent = rsS("Parent")
					Caccess = rsS("mnuAccess")
					Cfunct = rsS("mnuFunction")
			'response.Write("CaddMenu3: " & CaddMenu & "<br>")
			'response.Write("Ciname3: " & Ciname & "<br><br>")
		      		if CaddMenu <> "" then
					  if p_iTemplate = 5 then
			    	    writeNavHeaderV CaddMenu,mnuLink,ConClick,CaddMenu,Caccess,Cfunct,mx
					  else
			    	    writeNavHeaderH CaddMenu,mnuLink,ConClick,CaddMenu,Caccess,Cfunct,mx
					  end if
		  			  mx = mx + 1
			  		else
					  if p_iTemplate = 5 then
			    	    writeNavHeaderV mnuName,mnuLink,ConClick,Ciname,Caccess,Cfunct,mx
					  else
			    	    writeNavHeaderH mnuName,mnuLink,ConClick,Ciname,Caccess,Cfunct,mx
					  end if
		  			  mx = mx + 1
			  		end if
			  		rsS.movenext
				  loop
		  		end if
		  		set rsS = nothing
		  
			  else
					  if p_iTemplate = 5 then
			    	    writeNavHeaderV mnuName,mnuLink,ConClick,Ciname,Caccess,Cfunct,mx
					  else
			    		writeNavHeaderH mnuName,mnuLink,ConClick,Ciname,Caccess,Cfunct,mx
					  end if
		  		mx = mx + 1
			  end if
			  rsSb.movenext
			loop
		  end if
		  set rsSb = nothing
		else
					  if p_iTemplate = 5 then
			    	    writeNavHeaderV mnuName,mnuLink,ConClick,Ciname,Caccess,Cfunct,mx
					  else
		  				writeNavHeaderH mnuName,mnuLink,ConClick,Ciname,Caccess,Cfunct,mx
					  end if
		  mx = mx + 1
		  
		end if
		  rsMnuTop.movenext
	      mnuTree = mnuTree & appEnd(CappID)
	  loop
	  'next
		mnuTree = mnuTree & "</div>" & vbCRLF & vbCRLF
		
	  writeMenu(mnuTree)
		
	  for sm = 0 to ubound(arHeadRow)
		if trim(arHeadRow(sm,0)) <> "" then
		  H_MenuTree arHeadRow(sm,0),arHeadRow(sm,1),arHeadRow(sm,2)
		end if
	  next
	end sub
	
	Private Function writeNavHeaderV(mName,mLink,mOClick,mAM,mAcc,mFct,c)
		  mnuTree = mnuTree & accessStart(mAcc)
		  if mLink <> "" then
		    mnuTree = mnuTree & "<a href=""" & mLink & """ class=""menuButtonV"">" & mName & ""
			mnuTree = mnuTree & setMnuFunction(mFct)
			mnuTree = mnuTree & "</a>"
		  elseif mOClick <> "" then
		    mnuTree = mnuTree & "<a href=""javascript:;"" class=""menuButtonV"" onclick=""" & replace(mOClick,"''","'") & """>" & mName & ""
			mnuTree = mnuTree & setMnuFunction(mFct)
			mnuTree = mnuTree & "</a>"
		  else
			'lnk = lnk + 1
   		    Randomize()
  		    p_lnk=Int(Rnd()*9000)
			mnuTree = mnuTree & "<a href=""javascript:;"" class=""menuButtonV"" onmouseover=""buttonMouseover2(event, 'sub_" & p_lnk & "','vfly_menu');""><img src=""images/tri.gif"" align=""right"" class=""menuItemArrow"" border=""0"" vspace=""3"">" & mName & ""
			mnuTree = mnuTree & setMnuFunction(mFct)
			mnuTree = mnuTree & "</a>"
			
		    'ub = ubound(arHeadRow)+1
			'redim preserve arHeadRow(ubound(arHeadRow)+1,3)
			'response.Write("preserved: " & ubound(arHeadRow) & "<br>")
			'arHeadRow(mx,0) = t
			arHeadRow(c,0) = mName
			arHeadRow(c,1) = "sub_" & p_lnk
			arHeadRow(c,2) = mAM
		  end if
	   	  mnuTree = mnuTree & accessEnd(mAcc)
			'mnuTree = mnuTree & spcr
	End Function
	
	Private Function writeNavHeaderH(mName,mLink,mOClick,mAM,mAcc,mFct,c)
		  mnuTree = mnuTree & accessStart(mAcc)
		  if mLink <> "" then
		    mnuTree = mnuTree & "<a href=""" & mLink & """ class=""menuButton"">" & mName & "</a>" & vbCRLF
		  elseif mOClick <> "" then
		    mnuTree = mnuTree & "<a href=""javascript:;"" onclick=""" & replace(mOClick,"''","'") & """ class=""menuButton"">" & mName & "</a>" & vbCRLF
		  else
			'lnk = lnk + 1
   		    Randomize()
  		    p_lnk=Int(Rnd()*9000)
			mnuTree = mnuTree & "<a href=""javascript:;"" class=""menuButton"" onmouseover=""buttonMouseover(event, 'sub_" & p_lnk & "');"">" & mName & "</a>" & vbCRLF
			
		    'ub = ubound(arHeadRow)+1
			'redim preserve arHeadRow(ubound(arHeadRow)+1,3)
			'response.Write("preserved: " & ubound(arHeadRow) & "<br>")
			'arHeadRow(mx,0) = t
			arHeadRow(c,0) = mName
			arHeadRow(c,1) = "sub_" & p_lnk
			arHeadRow(c,2) = mAM
		  end if
	   	  mnuTree = mnuTree & accessEnd(mAcc)
	End Function
		
	Private function H_MenuTree(parent,mnID,iNam)
	  mnuSubTree = ""
	  sSQL = "SELECT * from Menu Where Parent ='" & parent & "' and iName = '" & iNam & "' order by mnuOrder asc"
	  Set rsSub = Server.CreateObject("ADODB.Recordset")
  	  rsSub.Open sSQL, my_Conn, 3, 1, &H0001
	  if not rsSub.eof then
  	    rSCount = rsSub.recordcount
		'redim arSubTree(rSCount,2)
        mnuTree = "<div id=""" & mnID & """ class=""" & fly_menu & """ onmouseover=""menuMouseover(event)"">" & vbCRLF
	    'do while not rsSub.eof 
		for mx = 0 to (rSCount - 1)
		  mnuName = rsSub("Name")
		  mnuLink = rsSub("Link")
		  ConClick = rsSub("onClick")
		  m_funct = rsSub("mnuFunction")
		  m_targ = rsSub("Target")
		  mnuTree = mnuTree & accessStart(rsSub("mnuAccess"))
			if mnuLink <> "" then
			    mnuTree = mnuTree & "<a href=""" & mnuLink & """ class=""" & fly_menu & "Item"" target=""" & m_targ & """>" & mnuName & ""
		      mnuTree = mnuTree & setMnuFunction(m_funct)
			  mnuTree = mnuTree & "<br></a>" & vbCRLF
			elseif ConClick <> "" then
			    mnuTree = mnuTree & "<a href=""javascript:;"" onClick=""" & replace(ConClick,"''","'") & """ class=""" & fly_menu & "Item"">" & mnuName & ""
		      mnuTree = mnuTree & setMnuFunction(m_funct)
			  mnuTree = mnuTree & "<br></a>" & vbCRLF
			else
   		      Randomize()
  		      p_lnk=Int(Rnd()*9000)
			  mnuTree = mnuTree & "<a class=""" & fly_menu & "Item"" href=""javascript:;"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'sub_" & p_lnk & "');""><span class=""menuItemText"">" & mnuName & "</span>"
		      mnuTree = mnuTree & setMnuFunction(m_funct)
			  mnuTree = mnuTree & "<span class=""menuItemArrow"">&#9654;</span><br></a>" & vbCRLF
			  mnuSubTree = mnuSubTree & mnuName & ":sub_" & p_lnk & ":" & rsSub("iName") & "|"
			end if
	   	  mnuTree = mnuTree & accessEnd(rsSub("mnuAccess"))
		  rsSub.movenext
		next
        mnuTree = mnuTree & "</div>" & vbCRLF & vbCRLF
		writeMenu(mnuTree)
		
		if mnuSubTree <> "" then
		  if instr(mnuSubTree,"|") <> 0 then
		    aTree = split(mnuSubTree,"|")
			for xm = 0 to ubound(aTree)-1
			  'response.Write(aTree(xm))
			  bTree = split(aTree(xm),":")
			  H_MenuTree bTree(0),bTree(1),bTree(2)
			next
		  else
		  end if
		end if
	  else
	    'response.Write("Menu not found")
	  end if
	  set rsSub = nothing
	end function
		
	
	Private Function writeMenu(s)
	  if p_createFile = 1 then
        Call Write2File(mnuFile, s)
	    'response.Write(execute(s))
		include.writeSource(s)
	  else
	  if FSOenabled then
	    include(mnuFile)
	  else
	    'response.Write(execute(s))
		include.writeSource(s)
	  end if
	  end if
	End Function 

Private Function DExists(d) 'true if file exists
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  DExists = fso.FolderExists(d)
  Set fso = Nothing
End Function
  
Private Function FExists(d) 'true if file exists
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  FExists = fso.FileExists(d)
  Set fso = Nothing
End Function
  
Private Function DelFile(f)
  If Trim(f)="" Then Exit Function  
	    'response.Write("<br>DelFile: " & f)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  if FExists(f) then fso.DeleteFile(f)
  Set fso = Nothing
End Function

Private Function FolderCount(dir)
  If Trim(dir)="" Then Exit Function  
  Dim fs
  Set fs = Createobject("Scripting.FileSystemobject") 
  Dim oFolder
  Set oFolder = fs.GetFolder(dir)
  FolderCount = oFolder.Files.Count  
  Set fs = Nothing
  Set oFolder = Nothing  
END Function

Public Function DelMenuFiles(f)
  'If Trim(f)="" Then Exit Function  
  Dim fs, oFolder
  Set fs = Createobject("Scripting.FileSystemobject") 
  mf = server.MapPath("files/config/menu")
  'p = fs.GetParentFolderName(mf)
  Set oFolder = fs.GetFolder(mf)
  for each i in oFolder.files
	  'DelFile(i.path)
	  fs.DeleteFile(i.path)
  next  
  Set oFolder = Nothing
  Set fs = Nothing
END Function

Public Function Write2File(afile,bstr)
  Dim wObj, wText, p, cf
  if afile="" Then EXIT Function
  if instr(afile,":") = 0 then afile = server.mappath(afile)
  Set wObj = CreateObject("Scripting.FileSystemObject")
  p=wObj.GetParentFolderName(afile)
  if DExists(p) then
	    'response.Write("<br>afile: " & afile)
    Set wtext = wObj.OpenTextFile(afile, 8, True)

    Dim nCharPos, sChar
    For nCharPos = 1 To Len(bstr)
      sChar = Mid(bstr, nCharPos, 1)
      On Error resume next  '<-- **** Error handing starts ****
      wtext.Write sChar
      On Error Goto 0       '<-- ***** Error handing ends *****
    Next

    wtext.Close()
    Set wtext = Nothing
	if FExists(afile) then
	    'response.Write("<br>File Exists<br>")
	else
	    'response.Write("<br>File Not Created<br>")
	end if
  else
    cf=wObj.GetParentFolderName(p)
	if DExists(cf) then
	  wObj.CreateFolder(p)
	  if DExists(p) then
	    Write2File afile,bstr
	  else
	    'Cannot create folder
	  end if
	else
	  wObj.CreateFolder(cf)
	  if DExists(cf) then
	    Write2File afile,bstr
	  else
	    'Cannot create folder
	  end if
	end if
  end if
  Set wObj = Nothing
End Function

Private Function ReadFileByLine(fpath)
  Dim fObj, ftext, fileStr  
  if fpath <> "" then
    if instr(fpath,":") = 0 then fpath = server.mappath(fpath)
  	Set fObj = CreateObject("Scripting.FileSystemObject")
  	If fObj.FileExists(fpath) Then
   	  Set ftext = fObj.OpenTextFile(fpath, 1, FALSE)
      fileStr =""
      WHILE NOT ftext.AtEndOfStream
      	fileStr  = fileStr  & ftext.ReadLine & chr(13)
      WEND
      ftext.Close
  	else
      fileStr = ""
  	End if
  End if
  ReadFile= fileStr
End Function

    private function ReadFile(str_path)
      dim objfso, objfile
      if str_path <> "" then
        if instr(str_path,":") = 0 then str_path = server.mappath(str_path)
        set objfso = server.createobject("scripting.filesystemobject")
        if objfso.fileexists(str_path) then
          set objfile = objfso.opentextfile(str_path, 1, false)
          if err.number = 0 then
            readfile = objfile.readall
            objfile.close
          end if
          set objfile = nothing
        end if
        set objfso = nothing
      end if
    end function
	

'****************************' Terminate Class	'*****************************
	Private Sub Class_Terminate()
	End Sub
	
End Class

dim mnu
set mnu = New SkyMenu
%>