<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Edit Code  - Processor</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	'strCodename = Request.QueryString("cd")
	'strCodeValue = Request.QueryString("cv")
	'strErrMsg    = Request.QueryString("ErrMsg")

	strCodeValue = Request.Form("CodeVal")
	strCodename  = Request.Form("CodeName")
response.write("strCodeName=>"&strCodeName&"<<br>")
	response.write("strErrMsg  =>"&strErrMsg&"<<br>")
if len(trim(strCodeName)) = 0 then
   strEM="The Codename is Blank"
   response.write(strEM)
   response.end
   response.redirect("AddCode0.asp?ErrMsg="&sstrEM)	
end if
	'*
	'**  Condition the input fields to prevent SQL Injection
	'*

	%> 
</body>
</html>
