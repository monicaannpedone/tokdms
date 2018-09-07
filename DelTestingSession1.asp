<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>pullData</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>

</head>

<body>

<%
   if Session("IsSysadmin") then
      intTestingID = Response.QueryString("tid")
      strSQL = "DELETE FROM dms_TestPull WHERE TestingID = "&intTestingID&";"
      oconn.execute strSQL
      strSQL = "DELETE FROM dms_TestSession WHERE TestingID = "&intTestingID&";"
      oconn.execute strSQL
      strSQL = "DELETE FROM dms_Testgroup WHERE TestingID = "&intTestingID&";"
      oconn.execute strSQL
   end if    
   response.redirect("Testing.asp") 

  
  
%>		
        
	</body>
</html>
