<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/md5.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validateSysAdmin.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Maintain Users</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
<div >
   <div >
	  <br/><br/>
		
<% 
Call TopLine()
strNewUserName = Request.Form("u")

strSQL = "SELECT * FROM dms_user WHERE UserName = '"&strNewUserName&"';"
oRS.Open strSQL

if NOT oRS.EOF then
   strMsg="User "&strNewUserName&" already exists."
   response.redirect("EditUser.asp?u="&strNewUserName&"&m="&strMode&"msg="&strMsg)
else
   ors.Close
   set ors = Nothing
end if

strNewUserName = Trim(Request.Form("Username"))
strFName    = Trim(Request.Form("FName"))
strLName    = Trim(Request.Form("LName"))
strEMail    = Trim(Request.Form("Email"))
strPhone    = Trim(Request.Form("UserPhone"))
if len(strPhone) > 0 then
	strPhone = replace(strPhone," ","")
	strPhone = replace(strPhone,"(","")
	strPhone = replace(strPhone,")","")
	strPhone = replace(strPhone,"-","")
	strPhone = replace(strPhone,".","")
end if
''response.write("IsActive = >"&Request.Form("IsActive")&"< <br/>")
'response.write("IsBlocked = >"&Request.Form("IsBlocked")&"< <br/>")
'response.write("Isadmin = >"&Request.Form("Isadmin")&"< <br/>")
'response.write("IsSysadmin = >"&Request.Form("IsSysadmin")&"< <br/>")

'response.write("Enddate = >"&Request.Form("EndDate")&"< <br/>")

datEndDate    = Request.Form("EndDate")
if datEnddate = "" then datEndDate=CDate("12/31/2099") end if
bitIsActive   = Request.Form("IsActive")
if bitIsActive = "" then bitIsActive = 0 else bitIsActive = 1 end if
bitIsBlocked  = Request.Form("IsBlocked")
if bitIsBlocked= "" then bitIsblocked = 0 else bitIsBlocked = 1 end if
bitIsAdmin    = Request.Form("IsAdmin")
if bitIsAdmin = "" then bitIsAdmin = 0 else bitIsAdmin = 1 end if
bitIsSysAdmin = Request.Form("IsSysAdmin")
if bitIsSysAdmin = "" then bitIsSysAdmin = 0 else bitIsSysAdmin = 1 end if
strDojoCode      = Request.Form("DojoCd")
bitCanChangeDojo = Request.Form("CanChangeDojo")
if len(bitCanChangeDojo) = 0 then  bitCanChangeDojo = 0 else bitCanChangeDojo = 1 end if
strDojoCd = strDojoCode

'
'' Edits
'
strErrmsg = ""
If len(strNewUserName) < 4 OR len(strNewUserName) > 16 then
   strErrMsg = "00Username must be between 4 and 16 characters."
end if
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
if IsDate(datEndDate) then
	if datEndDate < Date() then
	   strErrMsg = "05End date can not be in the past."
	end if
else
    strErrMsg="05End Date blank or invalid"	
end if	
'response.write("len(strErrMgs)="&len(strErrmsg)&"; strErrmsg=>"&strErrMsg&"< <br />")
if len(strErrmsg) > 0 then
'   response.end()
   response.redirect("AddUser.asp?err="&strErrMsg)
end if   

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
strCC = "postmaster@tokdms.com"
strText = "Dear "&strFname&" "&strLname&vbcrlf&vbcrlf&"A new account has been created for you."&vbcrlf
strText = strText&"UserName :"&strNewUserName&vbCRLF&"Password :"&strPassword&vbCRLF&vbCRLF
strText = strText &"The System Administrator has also been notified."
if NOT strLocale = "local" then
   intRC = DMSMail(strEmail,strCC, "New account created for user "&strNewUserName, strText)
end if
   response.write("</br>New account created for User: "&strNewUserName &";  Password: <b>"&strPassword&"</b></br>")
   response.write("Dojo: "&strDojoCd&"; CanChangeDojo:"&bitCanChangeDojo&"; </br>")
   response.write("Privileges Admin: "&bitIsAdmin&"; SysAdmin :"&bitIsSysAdmin&"</br></br>")
   strMsg = "User "&strNewUserName&" added by "&Session("UserName")
   Call TextSA(strMsg)
strPasswordEnc = md5(TRIM(strPassword))
'response.write("Password is "&strPassword&"<br />")
   strSQL = "INSERT INTO dms_user (Username, Password, PasswordEnc, Createdate, LastAccessdate, Enddate, IsActive, IsBlocked, IsAdmin, IsSysadmin  "
   strSQL = strSQL & ", Email, UserPhone, Userfname, userlname) "
   strSQL = strSQL & " VALUES ('"&strNewUserName&"','','"&strPasswordEnc&"','"&Now()& "','', '"&datEnddate&"', "&bitIsActive&", "&bitIsBlocked&" "
   strSQL = strSQL & ", "&bitIsAdmin&", "&bitIsSysadmin&" "
   strSQL = strSQL & ", '"&strEmail&"', '"&strPhone&"', '"&strFname&"', '"&strLName&"');"
'response.write("strSQL=>"&strSQL&"< <br />")
oconn.Execute strSQL
strSQL = "INSERT INTO dms_pref (Username, kaicd, dojocd, CanChangeDojo, SelActive, SelInactive, TestingID) "
strSQL = strSQL & "VALUES ('"&strNewUserName&"', 'abk', '"&strDojoCd&"', "&bitCanChangeDojo&", 1,1,0) ;"
'response.write("strSQL=>"&strSQL&"< <br />")
oconn.Execute strSQL
strAction = Session("UserName")&" Added User "&strNewUserName
Call ChangeJournal("dms_User", "username",  "addUser1", strAction)
strAction = Session("UserName")&" Added User preferences for "&strNewUserName
Call ChangeJournal("dms_pref", "username",  "addUser1", strAction)

'response.end()
%>
<br>
<br>
<% strUMsg = strNewUserName&" successfully added. " %>
<a href="maintusers.asp?msg=<%=strUMsg%>">Return to SysAdmin</a>  
<br>
		</div>
	</div>
</body>
</html>
