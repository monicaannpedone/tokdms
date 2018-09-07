<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validateSysAdmin.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!--#include file="../include/common.inc"-->

<head>
	<title>Delete User</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>
<img src="../images/tcslogo1.gif" />
<div >
   <div >
	  <br/><br/>
		
<% 
strNewUserName = Trim(Request.QueryString("u"))
strMode     = Trim(Ucase(Request.QueryString("m")))

strSQL = "SELECT * FROM dms_user WHERE UserName = '"&strNewUserName&"';"
oRS.Open strSQL

if oRS.EOF then
   strMsg="User >"&strNewUserName&"< Does Not Exist"
   response.redirect("MaintUsers.asp?msg="&strMsg)
else
   ors.Close
   set ors = Nothing
end if

strSQL = "DELETE FROM dms_user WHERE username = '"&strNewUserName&"'; "


oconn.Execute strSQL
strSQL = "DELETE FROM dms_pref WHERE UserName = '"&strNewUserName&"'; "
oconn.Execute strSQL

strMsg = "User "&strNewUserName&" Deleted by "&Session("UserName")
Call TextSA(strMsg)
strAction = "Deleted User "&strNewUserName
Call ChangeJournal("dms_User", "username",  "DeleteUser.asp", strAction)
strAction = "Deleted User preferences for "&strNewUserName
Call ChangeJournal("dms_pref", "username",  "DeleteUser.asp", strAction)

response.redirect("MaintUsers.asp?msg="&strMsg)

%> 
		</div>
        <!--#include file="../include/infooter.inc"-->  
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
