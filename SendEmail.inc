<%
EmailSubject = "Sending Email by CDO --  Test "&Date&" "&Time()
EmailBody = "This is the body of a message sent via" & vbCRLF & _
        "a CDO.Message object using SMTP authentication ,with port 465."

Const EmailFrom = "darkeraven@gmail.com"
Const EmailFromName = "Monica "
Const EmailTo = "darkeraven@gmail.com"
Const SMTPServer = "smtp.gmail.com"
Const SMTPLogon = "darkeraven@gmail.com"
Const SMTPPassword = "Sunsu1153!"
Const SMTPSSL = True
Const SMTPPort = 465

Const cdoSendUsingPickup = 1    'Send message using local SMTP service pickup directory.
Const cdoSendUsingPort = 2  'Send the message using SMTP over TCP/IP networking.

Const cdoAnonymous = 0  ' No authentication
Const cdoBasic = 1  ' BASIC clear text authentication
Const cdoNTLM = 2   ' NTLM, Microsoft proprietary authentication

' First, create the message

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = EmailSubject
objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
objMessage.To = EmailTo
objMessage.TextBody = EmailBody

' Second, configure the server

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update
'Now send the message!
On Error Resume Next
objMessage.Send

If Err.Number <> 0 Then
    MsgBox Err.Description,16,"Error Sending Mail"
Else 
    MsgBox "Mail was successfully sent !",64,"Information"
End If
%>