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
   strSQLDel  = "DELETE FROM dms_TestPullArchive;"
   oconn.execute strsqlDel
   strUserName = Session("Username")
   strSQLX = "SELECT * from dms_Testing order by TestDate, TestLocationID;"
   orsX.Open strSQLX
  
   if orsX.EOF then
      response.write("Testing File is Empty<br>")
      response.end
   End if
   intTestingID          = 1
   intRecCnt             = 0
   intLastTestingID      = -1
   datLastTestDate       = 0
   intLastTestLocationID = -1
   intTIDNo              = 0
   intRecsWritten        = 0
   intFlushCnt           = 0
   intTestingCount = 0
   intNumRecords = orsX.RecordCount
   do while NOT orsX.EOF and intRecCnt < 150000000
      intFlushCnt       = intFlushCnt + 1
      intTestingID      = orsX("TestingID")
      datTestDate       = FormatdateTime(orsX("TestDate"),2)
	  intTestLocationID = orsX("TestLocationID")
	  intStatusID       = orsX("StatusID")
	  intStudentID      = orsX("StudentID")
	  intRankID         = orsX("RankID")
      if intFlushCnt > 100 then
	     response.flush
		 response.write("<script type='text/javascript'>window.reload();</script>")
		 response.write("Processing Record "& intRecCnt&" of "&intNumRecords&"<br>")
		 intFlushCnt = 1
	  end if
      if NOT IsNull(datTestDate) then 
	     ' response.write("TestDate "&datLasttestDate&"::"&datTestdate&"  -  Loc "&intLastTestLocationID&"::"&intTestLocationID&"<br>")
	      if NOT (datLastTestDate = datTestDate AND intLastTestLocationID = intTestLocationID) then
		     
		     if intRecCnt > 1  then
			    
			    response.write("<td colspan='17' class='dtabh'><b> "&intTestingCount&" Records in Test Session <b></td>")
			    response.write("</table>") 
			 end if
			   
			
		     intTestingCount = 0
		     intTIDNO = intTIDNo + 1
   	         datTestdate = FormatDateTime(datTestdate,2)
	         if IsNull(intTestLocationID) then intTestLocationID = 2 end if
		 '    response.write("Rec"&intRecCnt&" TestDate="&dattesDate&"; TestLocationID="&intTestLocationID&"; <br>")
		     strTestingStatus = "6"
		     strSSDojoCd = "npk"
		     strTestingSchools = "11111"
		     strSQL =          "INSERT INTO dms_TestingSession (TestingID, SponsoringSchoolCd, TestingSchools, TestLocationCd, TestingDate,TestingStatus, LastChangeDate, TestIsActive, UpdatedBy) "
		     strSQL = strSQL & " VALUES ("&intTestingID&", '"&strSSDojoCd&"', '"&strTestingSchools&"', "&intTestLocationID&", '"&datTestDate&"', '"&strTestingStatus&"', "&Date()&", 1, 'mspedone') ;"
		    ' response.write("strSQL=>"&strSQL&"<<br>")
			 %><table>
			       <tr><td colspan='5'>&nbsp;</td></tr>
			       <tr>
				       <td><b>*** New Test Session</b> <%=intTIDNo%></td>
					   <td><b>Test Date</b> <%=datTestDate%></td>
					   <td><b>Location</b> <%= intTestLocationID %></td>
                       <td><b>Sponsor</b> <%=strSSDojoCd%></td>
                       <td><b>Schools</b> <%=strTestingSchools%></td>		   
				   </tr>
				   <tr><td colspan='5'>Test Group Created</td></tr>
				   <tr><td colspan='5'>Test SubGroup Created</td></tr>
			   </table>
			   <table>
			       <tr>
				       <td class="dtabh">&nbsp;</td>
					   <td class="dtabh">Test No</td>
					   <td class="dtabh">Student</td>
                       <td class="dtabh">Sponsor</td>
					   <td class="dtabh">Dojo </td>
					   <td class="dtabh">Status</td>
					   <td class="dtabh"> ==> </td>
					   <td class="dtabh">DojoCd </td>
	                   <td class="dtabh">RankID </td>
                       <td class="dtabh">Last Testing Date </td>
	                   <td class="dtabh">ToRankID  </td>
					   <td class="dtabh">Div </td>
					   <td class="dtabh">Anchor Day </td>
					   <td class="dtabh">Anchor Order </td>
					   <td class="dtabh">Status</td>
					   <td class="dtabh">Add Type</td>
					   <td class="dtabh">ExDojoCd </td>
				   </tr>
	   
			 <%
		  end if
		  datLastTestdate       = datTestDate 
		  intLastTestLocationID = intTestLocationID
	      intSID = intStudentID
			 ' *** Look up these values ***
			 Select Case(intTestLocationID)
			        Case(0)
					    strDojoCd = "npk"
					Case(1)
					    strDojoCd = "npk"
					Case(2)
					    strDojoCd = "npk"
					Case(3)
					    strDojoCd = "kin"
					Case(4)
					    strDojoCd = "bk"
					Case(5)
					    strDojoCd = "kin"
					Case(6)
					    strDojoCd = "bk"
					Case(7)
					    strDojoCd = "bk"
					Case(8)
						strDojoCd = "kin"
					Case(9)
					    strDojoCd = "npk"
					Case(10)
					    strDojoCd = "bk"
					Case(11)
					    strDojoCd = "kin"
					Case(12)
					    strDojoCd = "npk"
					Case(13)
					    strDojoCd = "npk"
					Case Other
					    strDojoCd = "npk"
			 End Select
			 intNextRankID      = intRankID
			
			 strAnchorDay       = "X"
			 intAnchordayOrder  = 8
			
			 intAddType         = 1
			 strXDojoCd         = ""
			 datLasttestingdate = datTestDate
			 strSSDojoCd = strDojoCd
			 
			 ' *** Look up these values
			
			 strSQLS = "SELECT StudentID, DojoCd, RankCd, Division, Anchorday FROM dms_Student WHERE StudentID = '"&intSID&"';" 
		'	 orsS.Open strSQLS
			' if orsS.EOF then
			'	response.write("Student >"&intSID&"< not found "&strSQLS&"<br>")
			' else
		'		strDojoCd    = orsS("DojoCd")
		'		intRankCd    = orsS("RankCd")
			'	strDivision  = orsS("Division")
			'	strAnchorday = Trim(orsS("Anchorday"))
			'	response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SID="&intSID&"; DojoCd="&strDojoCD&"; RankCd="&intRankCd&"; Division="&strDivision&"; AnchorDay="&strAnchorDay&";<br>")
			' end if		 
			' orsS.close
			' set orsS = Nothing
		'	 Set oRSS = Server.CreateObject("ADODB.Recordset") 
			' oRSS.ActiveConnection = oConn 


			 strSQLI = "INSERT INTO dms_testPullArchive ( TestingID, StudentID, DojoCd, RankID, LastTestingDate, NextRankID, Division, Anchorday, AnchordayOrder, StatusID, AddType, XDojoCd) "
			 strSQLI = strSQLI & "VALUES ("&intTestingID&", "&intSID&", '"&strDojoCd&"', "&intRankID&", '"&datLastTestingDate&"', "&intRankID&", '"&strDivision&"', '"&strAnchorday&"', "&intAnchordayOrder&","& intStatusID  &",'1','');"
			' response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;strSQLI=>"&strSQLI&"<<br>")
			 oConn.Execute strSQLI
			 %> 
			       <tr>
				       <td>&nbsp;</td>
					   <td><%=intTestingID%></td>
					   <td><%=datTestDate%></td>
                       <td><%=strSSDojoCd%></td>
					   <td><%= intTestLocationID %></td>
					   <td><%=intStatusID%></td>
					   <td>&nbsp;</td>
					   <td> <%=strDojoCd%></td>
	                   <td> <%=intRankID%></td>
                       <td><%=datLastTestingDate%></td>
	                   <td><%=intRankID%></td>
					   <td><%=strDivision%></td>
					   <td><%=strAnchorday%></td>
					   <td><%=intAnchordayOrder%></td>
					   <td><%=intStatusID%></td>
					   <td>1</td>
					   <td>&nbsp;</td>
				   </tr>
			    
			 <%
			 
	'	  set orsZ = Nothing
	'	  Set oRSZ = Server.CreateObject("ADODB.Recordset") 
		'  oRSZ.ActiveConnection = oConn 
		  intRecsWritten = intRecsWritten + 1
	  else 
           Response.write("@@@@@@@@@@ "&intRecCnt&"Test Date is Null, skipping the record<br>")	  
	  end if
      orsX.MoveNext	
	  intTestingID = intTestingID + 1
	  intRecCnt    = intRecCnt + 1
	  intTestingCount =  intTestingCount + 1
   Loop
   			    response.write("<td colspan='17' class='dtabh'><b> "&intTestingCount&" Records in Test Session <b></td>")
			    response.write("</table><br><br>") 

   response.write("<b>Records Read         </b>"&intRecCnt&"<br>")
   response.write("<b>Test Groups Created  </b>"&intTIDNo&"<br>")
   response.write("<b>Test Records Created </b>"&intRecsWritten&"<br>")
   
   orsX.Close
   set orsX = Nothing
   Set oRSX = Server.CreateObject("ADODB.Recordset") 
   oRSX.ActiveConnection = oConn 
   %>	
   <a href="converttesting1.asp">continue</a>

   
	
        <b>*** End***</b>
		</div>
	</div>
   </div>
	</body>
</html>
