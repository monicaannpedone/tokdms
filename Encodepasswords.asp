<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/md5.inc" -->
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Encode passwords</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>

<div >
   <div >
	  <br/><br/>
		
<% 
icnt = 0
strSQLX = "UPDATE dms_user SET PasswordEnc =  '';"
oconn.Execute strSQLX
response.write("Encoding Passords<br/>")
strSQL = "SELECT Username, Password, PasswordEnc FROM dms_user ;"
oRS.Open strSQL
do while NOT oRS.EOF
    icnt = icnt + 1
    strUsername     = oRS("UserName")
	strPassword     = trim(oRS("Password"))
	strPasswordEncx = oRS("PasswordEnc")
	strPasswordEnc  = md5(strPassword)
	strSQLU = "UPDATE dms_user SET PasswordEnc = '"&strPasswordEnc&"' WHERE Username = '"&strUsername&"';"
	oconn.Execute strSQLU
	ors.MoveNext		
loop
ors.Close
response.write("Records Processed:"& icnt&"; <br /><br />Review<br/><br />")
strSQL = "SELECT Username, Password, PasswordEnc FROM dms_user ;"
oRS.Open strSQL
icnt = 0
do while NOT oRS.EOF
    icnt = icnt + 1
    strUsername     = oRS("UserName")
	strPassword     = Trim(oRS("Password"))
	strPasswordEnc  = oRS("PasswordEnc")
    'response.write(icnt&"Userid: "&strUserName&"Password: "&strPassword&""&"EncPassword: "&strPasswordEnc&";</br>") 
	ors.MoveNext		
loop
ors.Close
'response.write("Records Processed:"& icnt&"; Completed<br />")
%> 
		</div>
	</div>
</body>
</html>
