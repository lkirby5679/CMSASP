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

select case lcase(strMailMode) 
	case "aspmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		If mTyp = 1  Then
		  objNewMail.ContentType = "text/html"
		Else
		  objNewMail.CharSet = 2
		End If

		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		'objNewMail.AddReplyTo = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.BodyText = str_Msg
		on error resume next '## Ignore Errors
		SendOk = objNewMail.SendMail
		If not(SendOk) <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & objNewMail.Response & "</li>"
		End if
		on error goto 0
	case "aspemail"
		Set objNewMail = Server.CreateObject("Persits.MailSender")
		objNewMail.FromName = strFromName
		objNewMail.From = strSender
		objNewMail.AddReplyTo strSender
		objNewMail.Host = strMailServer
		objNewMail.AddAddress strRecipients, strRecipientsName
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "aspqmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		objNewMail.QMessage = 1
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.BodyText = str_Msg
		on error resume next '## Ignore Errors
		objNewMail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "cdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		objNewMail.From = strSender
		objNewMail.To = strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg

		If mTyp = 1 Then
		  objNewMail.Bodyformat=0  
		  objNewMail.Mailformat=0
		else
		  objNewMail.BodyFormat = 1
		  objNewMail.MailFormat = 0
		End If
		if attachment = 1 then
            objNewMail.AttachFile strIcsLocation '## Un-Comment for CDONTS
        end if
		
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "chilicdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		on error resume next '## Ignore Errors
		objNewMail.Send strSender, strRecipients, str_Subj, str_Msg
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "cdosys"
	        Set objNewMail = Server.CreateObject("CDO.Message")
	        Set iConf = Server.CreateObject ("CDO.Configuration")
        	Set Flds = iConf.Fields 

	        'Set and update fields properties
        	Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
	        Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer
		'Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"
		'Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
		'Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		'Flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

        	Flds.Update
        	Set objNewMail.Configuration = iConf
	        'Format and send message
        	Err.Clear 
			objNewMail.From = strSender
			objNewMail.To = strRecipients
			'objNewMail.Bcc = strRecipients
			objNewMail.Subject = str_Subj
			if mTyp = 1 then
               objNewMail.HTMLBody = str_Msg
			else
               objNewMail.TextBody = str_Msg
			end if
        	On Error Resume Next
        
		if attachment = 1 then
        	objNewMail.AddAttachment strIcsLocation
        end if
		objNewMail.Send
		Set iConf = Nothing
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
			response.Write(Err_Msg & "<br>")
		End if
		on error goto 0
	case "dkqmail"
		Set objNewMail = Server.CreateObject("dkQmail.Qmail")
		objNewMail.FromEmail = strSender
		objNewMail.ToEmail = strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg
		objNewMail.CC = ""
		objNewMail.MessageType = "TEXT"
		on error resume next '## Ignore Errors
		objNewMail.SendMail()
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "geocel"
		set objNewMail = Server.CreateObject("Geocel.Mailer")
		objNewMail.AddServer strMailServer, 25
		objNewMail.AddRecipient strRecipients, strRecipientsName
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strFrom
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg
		on error resume next '##  Ignore Errors
		objNewMail.Send()
		if Err <> 0 then 
			Response.Write "" & txtEmlError & ": " & Err.Description 
		else
			Response.Write "Your mail has been sent..."
		end if
		on error goto 0
	case "iismail"
		Set objNewMail = Server.CreateObject("iismail.iismail.1")
		MailServer = strMailServer
		objNewMail.Server = strMailServer
		objNewMail.addRecipient(strRecipients)
		objNewMail.From = strSender
		objNewMail.Subject = str_Subj
		objNewMail.body = str_Msg
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "jmail"
		Set objNewMail = Server.CreateObject("Jmail.smtpmail")
		objNewMail.ServerAddress = strMailServer
		If mTyp = 1 THEN
		  objNewMail.ContentType = "text/html"
		End If
		objNewMail.AddRecipient strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = str_Subj
		objNewMail.body = str_Msg
		objNewMail.priority = 3
		on error resume next '## Ignore Errors
		objNewMail.execute
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
	case "smtp"
		Set objNewMail = Server.CreateObject("SmtpMail.SmtpMail.1")
		objNewMail.MailServer = strMailServer
		objNewMail.Recipients = strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = str_Subj
		objNewMail.Message = str_Msg
		on error resume next '## Ignore Errors
		objNewMail.SendMail2
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>" & txtEmlError & ": " & Err.Description & "</li>"
		End if
		on error goto 0
end select

Set objNewMail = Nothing
%>
