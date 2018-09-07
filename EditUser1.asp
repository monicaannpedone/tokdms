<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/md5.inc" -->


<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->

<head>
	<title>Edit User</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>
<div >
   <div >
	  <br/><br/>
		
<% 
strUserName = TRIM(Request.Form("u"))

strSQL = "SELECT * FROM dms_user WHERE UserName = '"&strUsername&"';"
oRS.Open strSQL

if oRS.EOF then
   strMsg="00User "&strUsername&"Does Not Exist"
   response.redirect("EditUser.asp?u="&strUsername&"&m="&strMode&"msg="&strMsg)
end if
strOldEmail = oRS("Email")
oRS.Close
set oRS = Nothing
strFName    = Trim(Request.Form("FName"))
strLName    = Trim(Request.Form("LName"))
strEMail    = Trim(Request.Form("Email"))
strPhone    = Trim(Request.Form("Phone"))
if len(strPhone) > 0 then
	strPhone = replace(strPhone," ","")
	strPhone = replace(strPhone,"(","")
	strPhone = replace(strPhone,")","")
	strPhone = replace(strPhone,"-","")
	strPhone = replace(strPhone,".","")
end if
datEndDate    = Request.Form("EndDate")
bitresetPW   = Request.Form("ResetPW")
if bitResetPW = "" then bitResetPW = 0 else bitResetPW = 1 end if
if datEnddate = "" then datEndDate=CDate("5/12/2012") end if
bitIsActive   = Request.Form("IsActive")
if bitIsActive = "" then bitIsActive = 0 else bitIsActive = 1 end if
bitIsBlocked  = Request.Form("IsBlocked")
if bitIsBlocked= "" then bitIsblocked = 0 else bitIsBlocked = 1 end if
bitIsAdmin    = Request.Form("IsAdmin")
if bitIsAdmin = "" then bitIsAdmin = 0 else bitIsAdmin = 1 end if
bitIsSysAdmin = Request.Form("IsSysAdmin")
if bitIsSysAdmin = "" then bitIsSysAdmin = 0 else bitIsSysAdmin = 1 end if

strComment    = Trim(Request.Form("Comment"))
'
'' Edits
'
strErrmsg = ""
If len(strFname) < 1  then
   strErrMsg = "01First name can not be blank."
end if
If len(strLname) < 1  then
   strErrMsg = "02Last name can not be blank."
end if
intPos = instr(strEmail,"@")
if intPos = 0 then
   strErrMsg = "03Email address invalid."
end if
'if IsDate(datEndDate) then
'	if datEndDate < Date() then
'	   strErrMsg = "05End date can not be in the past."
'	end if
'else
'    strErrMsg="05End Date blank or invalid"	
'end if	
'response.write("len(strErrMgs)="&len(strErrmsg)&"; strErrmsg=>"&strErrMsg&"< <br />")
if len(strErrmsg) > 0 then
'   response.end()
   response.redirect("EditUser.asp?err="&strErrMsg&"&u="&strusername)
end if   


strPW=""
if bitResetPW then
	strAlphabet="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890@#$"
	intLenAlphabet=len(strAlphabet)
	intlenPassword = 10
	strPassword = ""
	Randomize
	for intPasswordCharacters = 1 to intLenPassword
		fltRnd = Rnd()
		intRnd = int(fltRnd*intLenAlphabet+.5)
		strPassword = strPassword & mid(strAlphabet,intRnd+1,1)
	   ' response.write("ii="&ii&"; fltRnd="&fltRnd&"; intRnd="&intRnd&"; strPassword=>"&strPassword&"< <br/>")
	next
	strPasswordEnc = md5(strPassword)
	strPWE = ", PasswordEnc = '"&strPasswordEnc&"' "
	strPW  = ", Password = '' "
strCC = "postmaster@tokdms.com"
strText = "Dear "&strFname&" "&strLname&vbcrlf&vbcrlf&"A new account has been created for you."&vbcrlf
strText = strText&"UserName :"&strNewUserName&vbCRLF&"Password :"&strPassword&vbCRLF&vbCRLF
strText = strText &"The System Administrator has also been notified."
if NOT strLocale = "local" then
   intRC = DMSMail(strEmail,strCC, "New account created for user "&strUserName, strText)
end if
   response.write("</br>Account Changed for User: "&strUserName &";  Password: <b>"&strPassword&"</b></br>")
   response.write("Dojo: "&strDojoCd&"; CanChangeDojo:"&bitCanChangeDojo&"; </br>")
   response.write("Privileges Admin: "&bitIsAdmin&"; SysAdmin :"&bitIsSysAdmin&"</br></br>")
   strMsg = "User "&strUserName&" added by "&Session("UserName")
   Call TextSA(strMsg)
   response.write("Password changed to "&strPassword)
end if



strSQL = "UPDATE dms_user SET Enddate='"&datEndDate&"', IsActive="&bitIsActive&", IsBlocked="&bitIsBlocked&"  "&strPW&strPWE
strSQL = strSQL & ", Email='"&strEmail&"', UserPhone='"&strPhone&"', Userfname='"&strFname&"', userlname='"&strLName&"' "
strSQL = strSQL & ", comment='"&strComment&"',  IsAdmin="&bitIsAdmin&", IsSysadmin = "&bitIsSysAdmin&" "
strSQL = strSQL & " WHERE username = '"&strUsername&"'; "


'response.write("strSQL=>"&strSQL&"< <br />")
oconn.Execute strSQL
strAction = Session("UserName")&" Changed User "&strUserName
Call ChangeJournal("dms_User", "username",  "editUser1", strAction)
%>
<br>
<br>
<% strUMsg = strUserName&" successfully Changed. " %>
<a href="maintusers.asp?msg=<%=strUMsg%>">Return to SysAdmin</a>  
<br>



		</div>
	</div>
</body>
</html>
