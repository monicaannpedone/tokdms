<%
Function DMSMail(strTo,strCC,strTitle, strText)
 
Set Mailer = Server.CreateObject("Persits.MailSender") 
 'response.write("To:>"&strTo&"<</br>")
 'response.write("CC:>"&strCC&"<</br>")
 'response.write("Title:>"&strTitle&"<</br>")
 'response.write("Text:>"&strText&"<</br>")
Mailer.Host = "mail.tokdms.com" 
Mailer.Username = "postmaster@tokdms.com" 
Mailer.Password = "Sanchin1153!" 
Mailer.FromName   = "postmaster@tokdms.com" 
Mailer.From= "postmaster@tokdms.com" 
Mailer.AddAddress strTo
Mailer.AddCC  strCC
Mailer.Subject    = strTitle
Mailer.Body   = strText 
 
' catch errors 
On Error Resume Next 
 
Mailer.Send 
 
If Err <> 0 Then ' error occurred 
      set TCSMail = Err 
else 
     'Response.Write "Mail sent..." 
	 set TCSMail = 0
End If 
Set Mailer = Nothing
 
End Function
%>