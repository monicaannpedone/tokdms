<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Delete Code Processor</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
	
	<%
		strCodeName      = Request.Form("CodeName")
        strCodeValue     = Request.Form("CodeValue")
        Response.write("strCodename=>"&strCodename&"<<br>")

        Response.write("Request.Form(CodeName)=>"&Request.Form("CodeName")&"<<br>")
        Response.write("strCodeValue=>"&strCodeValue&"<<br>")
        

	  
	strSQLI = "DELETE FROM dms_codes WHERE CodeName='"&strCodeName&"' AND CodeValue='"&strCodeValue&"' ;"
	response.write("strSQLI=>"&strSQLI&"<<br>")
	
	'oConn.Execute(strSQLI)
			
	ors.Close
    set ors = Nothing
	response.redirect("EditCode.asp?cd="&strCodeName)
	%> 
</body>
</html>
