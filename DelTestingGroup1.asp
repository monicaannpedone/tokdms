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

<%    intTestingID = Request.QueryString("t")
      intTestGroup = Request.QueryString("g")
      intTestSubGroup = Request.QueryString("sg")
      strSQLGroup = "DELETE FROM dms_TestSubGroup WHERE TID="&intTestingID&" AND  TG="&intTestGroup&" AND TSG="&intTestSubGroup&";"
      oConn.Execute strSQLGroup
      strSQL= "SELECT TID, TG, TSG FROM dms_TestSubGroup WHERE  TID="&intTestingID&" AND  TG="&intTestGroup&";"
      orsC.Open strSQL
      if orsC.EOF then
         strSQLGroup = "DELETE FROM dms_TestGroup WHERE TestingID="&intTestingID&" AND  TestingGroup="&intTestGroup&";"
         oConn.Execute strSQLGroup
      end if
     
   response.redirect("SetGroupTimes.asp?t="&intTestingID) 

  
  
%>		
        
	</body>
</html>
