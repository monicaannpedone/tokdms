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

<%    intTestingID = Request.QueryString("tid")
      strSQL = "SELECT TestingSchools from dms_TestSession where TestingID = "&intTestingID&";"
      orsc.Open strSQL
      if NOT orsc.EOF then
         strTestingSchools = TRIM(orsC("TestingSchools"))
      end if
      orsC.Close
      set orsC = Nothing
      Set oRSC = Server.CreateObject("ADODB.Recordset") 
      oRSC.ActiveConnection = oConn 
     
      strSQL = "Select * FROM dms_TestGroup WHERE TestingID = "&intTestingID&" and TGIsActive = 1 ORDER BY TestingGroup;" 
      orsc.Open strSQL
      intNGroups = 0
      do while NOT orsC.EOF
         intNGroups = intNGroups + 1
         intTG = orsC("TestingGroup")
         orsC.Movenext
      loop 
      intNextTG = intTG + 1
      orsC.Close
     response.write("Number of Groups is "&intnGroups&", the next Group is "&intNextTG&"<br>")
  ' response.flush()
     ' response.end

      set orsC = Nothing
      Set oRSC = Server.CreateObject("ADODB.Recordset") 
      oRSC.ActiveConnection = oConn 
 
         strSQLGroup = "INSERT INTO dms_TestGroup   (TestingID, TestingGroup, TGSeq, ArrivalTime, TestTime,     UpdatedBy, UpdateDateTime, TGIsActive)"
         strSQLGroup = strSQLGroup &  " VALUES ("&intTestingID&",    "&intNextTG&", "&intNextTG&",0, 0, '"&strUsername&"', '"&Now()&"', 1 );"
         response.write("strSQLGroup=>"&strSQLGroup&"<<br>")
         oConn.Execute strSQLGroup
      

     strUserName = Session("UserName")
     intii = 0
     for inti = 1 to len(strTestingSchools)
        if mid(strTestingSchools,inti,1) = 1 then
           intii = intii + 1
           if inti = 1 then strDojoCd = "npk" end if
           if inti = 2 then strDojoCd = "bk " end if
           if inti = 3 then strDojoCd = "kin" end if
           if inti = 4 then strDojoCd = "pv " end if
           if inti = 5 then strDojoCd = "ef " end if
           
           strSQLSubGroup =    "INSERT INTO dms_TestSubGroup   (     TID, TG,       TSG,      DojoNum,          DojoCd, Division, FromRankID, ToRankID,       UpdatedBy, UpdateDateTime, TSGIsActive)"
           strSQLSubGroup = strSQLSubGroup &  " VALUES ("&intTestingID&",   "&intNextTG&", "&intii&",     "&inti&", '"&strDojoCd&"',   '1111',          1,       15,'"&strUsername&"',  '"&Now()&"', 1 );"
           response.write("strSQLSubGroup=>"&strSQLSubGroup&"<<br>")
           oConn.Execute strSQLSubGroup
        end if
     next   
     response.redirect("SetGroupTimes.asp?t="&intTestingID) 

%>		
        
	</body>
</html>
