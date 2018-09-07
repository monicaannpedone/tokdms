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
      intG         = Request.QueryString("g")
       strUserName = Session("UserName")

      response.write("TestingID=>"&intTestingID&"<; TestGroup=>"&intG&"<<br>")
     
      strSQL = "SELECT TestingSchools, SponsoringSchoolCd from dms_TestSession where TestingID = "&intTestingID&";"
      orsc.Open strSQL
      if NOT orsc.EOF then
         strTestingSchools = TRIM(orsC("TestingSchools"))
         strSponsoringSchoolCd  = Trim(orsC("SponsoringSchoolCd"))
      end if
      orsC.Close
      set orsC = Nothing
      Set oRSC = Server.CreateObject("ADODB.Recordset") 
      oRSC.ActiveConnection = oConn 
     
      strSQL = "SELECT MAX(TSG) AS MaxTSG FROM dms_TestSubGroup WHERE TID = "&intTestingID&" AND TG= "&intG&";" 
      orsc.Open strSQL
      if intnGroups < 129 then
         if orsC.EOF then
            intMaxTSG = 1
         else
            intMaxTSG = orsC("MaxTSG") + 1
         end if 
                   response.write("EOF="&orsc.EOF&", Max="&orsC("MaxTSG")&"intMAXTSG="&intMaxTSG&" <br>")

		 intii = 0
		 inti = 0
		 strDojoCd = strSponsoringSchoolCd  
		 if strDojoCd = "npk" then intDojoCd = 1 end if 
		 if strDojoCd = "bk"  then intDojoCd = 2 end if 
		 if strDojoCd = "kin" then intDojoCd = 3 end if 
		 if strDojoCd = "pv"  then intDojoCd = 4 end if 
		 if strDojoCd = "ef"  then intDojoCd = 5 end if 
           
         strSQLSubGroup =    "INSERT INTO dms_TestSubGroup   (     TID, TG,       TSG,      DojoNum,          DojoCd, Division, FromRankID, ToRankID,       UpdatedBy, UpdateDateTime, TSGIsActive)"
         strSQLSubGroup = strSQLSubGroup &  " VALUES ("&intTestingID&", "&intg&",  "&intMaxTSG&",      "&intDojoCd&", '"&strDojoCd&"',   '1111',          1,       15,'"&strUsername&"',  '"&Now()&"', 1 );"
         response.write("strSQLSubGroup=>"&strSQLSubGroup&"<<br>")
    
         oConn.Execute strSQLSubGroup
     end if

   response.redirect("SetGroupTimes.asp?t="&intTestingID) 

%>		
        
	</body>
</html>
