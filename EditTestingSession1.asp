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
      <% Call TopLine() %>
	  <div id="contents">
	  <div id="OneCol">

<%
   strTID            = Request.Form("tid")
   strSDojo          = Request.Form("SSC")
   strdn01           = Request.Form("dn1")
   strdn02           = Request.Form("dn2")
   strdn03           = Request.Form("dn3")
   strdn04           = Request.Form("dn4")
   strdn05           = Request.Form("dn5")
   strTestLocation   = Request.Form("TestLocation")
   datTestingDate    = Request.Form("TestingDate")
   strTestSteps      = Request.Form("TestSteps")
   bitTestIsActive   = Request.Form("TestIsActive")
   if bitTestIsActive = "" then bitTestIsActive = 0 end if
   response.write("Test Active1? >"&bitTestIsActive&"<<br>")
      strPDojoList = ""
      if strdn01 = "1"            then strPDojoList = strPDojoList & "1" else strPDojoList = strPDojoList & "0" end if
      if strdn02 = "1"            then strPDojoList = strPDojoList & "1" else strPDojoList = strPDojoList & "0" end if
      if strdn03 = "1"            then strPDojoList = strPDojoList & "1" else strPDojoList = strPDojoList & "0" end if
      if strdn04 = "1"            then strPDojoList = strPDojoList & "1" else strPDojoList = strPDojoList & "0" end if
      if strdn05 = "1"            then strPDojoList = strPDojoList & "1" else strPDojoList = strPDojoList & "0" end if
 
  
   strUserName = Session("Username")
   strSQL = "SELECT * FROM dms_pref Where username='"&strUserName&"';"
   orsX.Open strSQL
   if orsX.EOF then
      response.write("Userid >"&strUserName&"< not found<br>")
      response.end
   End if
   bitCanChangeDojo = orsX("CanChangeDojo")
   strDefaultDojo   = orsX("DojoCd")
   intTestingID     = orsX("TestingID")
   ' response.write("Username=>"&strUserName&"<, CanChangeDojo=>"&bitCanChangeDojo&"<, DojoCd=>"&strDojo&"<<br>")
   orsX.Close
   set orsX = Nothing
   Set oRSX = Server.CreateObject("ADODB.Recordset") 
   oRSX.ActiveConnection = oConn 
   
  response.write("tid=>"&strtid&"<<br>")
  response.write("Sponsoring Dojo=>"&strSDojo&"<<br>")
  response.write("1234567890<br>")
  response.write(strPDojoList&"<br>")

  response.write("Sponsoring    Dojo>"&strSDojo&"<<br>")
  response.write("Participating Dojo>"&strPDojo&"<<br>")
  response.write("Testing Location  >"&strTestLocation&"<<br>")
  response.write("Testing Date      >"&datTestingDate&"<<br>")
  response.write("Test Status       >"&strTestSteps&"<<br>")
  response.write("Test Active? >"&bitTestIsActive&"<<br>")
  response.write("Next Testing ID =>"&intTestingID&"<<br>")
  strSQL = "SELECT * FROM dms_TestSession WHERE TestingID = "&strTID&";"
  response.write("strSQL=>"&strSQL&"<<br>")
  orsZ.Open strSQL
  if NOT orsZ.EOF then
     strUserName = Session("UserName")
     strSQLI = "UPDATE dms_TestSession SET SponsoringSchoolCd = '"&strSDojo&"', TestingSchools= '"& strPDojoList&"', TestingLocationCd =  '"&strTestLocation&"', "
     strSQLI = strSQLI & " TestingStatus='"&strTestSteps&"',	 TestingDate = '"&datTestingDate&"', TestIsActive = "&bitTestIsActive&", LastChangedate= '"&date()&"',  UpdatedBy = '"&strUsername&"' ;"
     oConn.Execute strSQLI
     response.write("strSQLi=>"&strSQLi&"<<br>")
    
  else
     response.write("That testing ID, "&strTID&" does nor exist!!!")
     response.end	 
  end if
  
  response.redirect("EditTestingSession.asp?em="&strEM&"&tid="&strtid)
%>		
        
		</div>
	</div>
   </div>
	</body>
</html>
