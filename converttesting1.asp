<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>Convert Testing Info</title>
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
   strSQLStud = "SELECT StudentID, DojoCD, AnchorDay, Division, LName, FName from dms_Student ORDER BY StudentID;"
   strSQLTP   = "SELECT * FROM dms_testPullArchive ORDER BY StudentID, LastTestingDate;"
   orsX.Open strSQLTP
   if orsX.EOF then
      response.write("Testing File is EMPTY!!!<br>")
	  response.end
   end if
   orsD.Open strSQLStud
   if orsD.EOF then   
      response.write("Student File is EMPTY!!!<br>")
	  response.end
   end if
   response.write("&&&&& Updating Student &&&&&<br>")
   
   intSIDD = -1
   intReads = 0
   intMatches = 0
   intWrites = 0
   intFlushCnt = 0
   do while NOT orsX.EOF
      if intFlushCnt > 100 then
	     response.flush
		 response.write("<script type='text/javascript'>window.reload();</script>")
		 response.write("Processing Record "& intReads&" <br>")
		 intFlushCnt = 1
	  end if
      intFlushCnt = intFlushCnt + 1
      intSIDX      = orsX("StudentID")
	  intTestingID = orsX("TestingID")
	  response.write("Looking Up Student "&intSIDX&"<br>")
      do while (NOT orsD.EOF) AND intSIDD < intSIDX
		 orsD.MoveNext
		 if NOT orsD.EOF Then
	        intSIDD = orsD("StudentID")
	      '  response.write("&nbsp;&nbsp;&nbsp;  Getting Next Student"&intSIDX&"; Student "&intSIDD&"<br>")
	     else
		    response.write("<<<<<<<<<<<<<<<<<<<<<<<  EOF on Student >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><br>")
		 end if	
	  loop
	  intMatches = intMatches + 1
	  strLName          = orsD("LName")
	  strFName          = orsD("FName")
	  strDojoCd         = OrsD("DojoCd")
	  strAnchorday      = Trim(orsD("AnchorDay"))
	  response.write("Anchor Day "&strAnchorday&"<br>")
      SELECT Case(strAnchorDay)
	         Case ("M")
			      intAnchorDayOrder = 2
	         Case ("T")
			      intAnchorDayOrder = 3
	         Case ("W")
			      intAnchorDayOrder = 4
	         Case ("H")
			      intAnchorDayOrder = 5
	         Case ("F")
			      intAnchorDayOrder = 6 
	         Case ("A")
			      intAnchorDayOrder = 7
	         Case ("M")
			      intAnchorDayOrder = 8
	         Case ("S")
			      intAnchorDayOrder = 1
	         Case Else
			      intAnchorDayOrder = 8
	  END SELECT
	  strDivision       = orsD("Division")
	 ' response.write("&nbsp;&nbsp;&nbsp;&nbsp;"&intMatches&" Match Test Pull "&intSIDX&"; Student "&intSIDD&" "&strFName&" "&strLName&"<br>")
	  
	  intReads = intReads + 1
	  strSQLUpDate = "UPDATE dms_TestPullArchive SET DojoCd = '"&strDojoCd&"', Division = '"&strDivision&"', AnchorDay = '"&strAnchorDay&"', AnchorDayorder = "&intAnchorDayOrder&", XREF="&intTestingID
	  strSQLUpDate = strSQLUpdate & "  WHERE TestingID ="&intTestingID&" AND StudentID = "&intSIDX&" ;"
     ' response.write("UPDATE >"&strSQLUpDate&"<<br>")
	  oconn.Execute strSQLUpdate
	  orsX.MoveNext'
%>		
        <b>*** End***</b>
		</div>
	</div>
   </div>
	</body>
</html>
