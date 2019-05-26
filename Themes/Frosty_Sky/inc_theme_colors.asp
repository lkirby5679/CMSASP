<% 
thmFolder = "Skin"				        'Theme folder name
thmAuthor = "<A HREF=http://www.allartprints.net/>R.Frost</A>"	'Theme Author
thmDescription = "Frosty Sky"                'Theme Description
thmLogoImage = "Site_Logo.jpg"			        	'Site logo
thmSubSkin = 1	

Session("thmFolder") = thmFolder
Session("thmAuthor") = thmAuthor
Session("thmDescription") = thmDescription
Session("thmLogoImage") = thmLogoImage
Session("thmSubSkin") = thmSubSkin

newThm = replace(replace(request("tName"),"<",""),">","")
thmFolder = replace(replace(request("tFolder"),"<",""),">","")
whereto =  request.ServerVariables("HTTP_REFERER") & "?tName=" & newThm & "&tFolder=" & thmFolder & "&cmd=1"
'whereto = "admin_config_themes.asp?tName=" & newThm & "&tFolder=" & thmFolder
response.Redirect whereto
%>