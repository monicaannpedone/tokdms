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
<div id="main">
	  <div id="contents">
	  <div id="OneCol">

<%
    intTID    = Request.Querystring("tid")
    strSt     = Request.QueryString("st")
    response.write("intTid=>"&intTID&"<<br>")
    strUserName = Session("Username")
    strSQLU = "UPDATE dms_TestSession SET  TestingStatus='"&strSt&"', LastChangeDate='"&Date()&"', UpdatedBy='"&strUserName&"'  WHERE TestingID = "&intTID&";"
    oConn.Execute strSQLU

    strEM = "Test Session "&intTID&" set to Status "&strSt
    response.redirect("Testing.asp?msg="&strEM)
  
  
%>		
        
		</div>
	</div>
   </div>
	</body>
</html>
