<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>Edit Testing Session Data</title>
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
    <div id="titlecontents">
        <div id="TitleBox">
            <h2><a href="Admin.asp"><i class="fa fa-bars"></i></a>&nbsp;&nbsp;<img src="../images/lillilinvertedtomoe.gif" />&nbsp;Dojo Management System </h2>
        </div>   
        <div id="AcctBox">
            <h4><a href="admin.asp"><%=strUserName%>&nbsp;<i class="fa fa-user"></i></a>&nbsp;&nbsp;&nbsp;<a href="sysadmin.asp"><i class="fa fa-cog"></i></a>&nbsp;&nbsp;<a href="Logout.asp"><i class="fa fa-sign-out"></i></a></h4>
        </div>
    </div>  
	  <div id="contents">
	  <div id="OneCol">
	  <h1>Edit Testing Session Data</h1>
	  <h3></h3>
      <a href="Testing.asp">Return to Testing Menu</a><br><br>
	

<%
 DIM aryDojoCds(32)
 strSQLD = "SELECT DojoCd from dms_Dojo WHERE DojoIsActive = 1 ORDER BY DojoOrder"
 orsD.Open strSQLD
 intD = 0
 do while NOT orsD.EOF
    intD = intD + 1
    aryDojoCds(intD) = orsd("DojoCd")
    orsd.MoveNext
 loop
 orsD.Close
 
    strTID = Request.QueryString("tid")
    strSQL = "SELECT TestingID, SponsoringSchoolCd, TestingSchools, TestingLocationCd, TestingDate, TestingStatus, LastChangeDate, TestIsActive, UpdatedBY FROM dms_TestSession WHERE TestingID="&strTID&";"
    orsT.Open strSQL
	if orsT.EOF then
	   response.write("Invalid Test Session ID - Fatal Error")
	   response.end
	end if

    intTID                = orsT("TestingID")
    strSSC                = orsT("SponsoringSchoolCd")
    strTestingSchools     = orsT("TestingSchools")	
	strTestingLocationCD  = orsT("TestingLocationCd")
	datTestingDate        = orsT("TestingDate")
	datTestingDateMMDDYY  = FormatDateTime(datTestingDate,2)
	strTestingStatus      = Trim(orsT("TestingStatus"))
	datLastChangeDate     = orsT("LastChangeDate")
	bitTestIsActive       = orsT("TestIsActive")
	strUpdatedBy          = orsT("UpdatedBy")
	
		
	%> 
   <table>
      <form name="EditTS" action="EditTestingSession1.asp" method="post" id="EditTS">
      <tr> 
	     <td>Test Session</td>
		 <td><%=intTID %><input type="hidden" value="<%=strTID%>" name="tid"></td> 
		 <td>&nbsp;</td>
	   </tr>
	   <tr>
	     <td>Sponsoring School&nbsp;&nbsp;&nbsp;</td>
		 <td><% 
		     Call SelectDojo("ssc", strSSC, 0)
			 %>
         </td>
		 <td>&nbsp;</td>
		 
	   </tr>
	   <tr>
	     <td>Testing Schools </td>
		 <td>
            <%
			for inti = 1 to 5
				strVal = Mid(strTestingSchools,inti,1)
				strdn  = aryDojoCds(inti)
				strN   = "dn"&CStr(inti)
				Call SelectCheckboxLabel(strN,strVal,strdn)
            next 
			%>
			</td>
		 <td>&nbsp;</td>
			
	   </tr>
	   <tr>
	     <td>Testing Location</td>
		 <td><%
		     Call SelectFromTable("TestLocation",strTestingLocationCd)
	         %>
		 </td>
		 <td>&nbsp;</td>
		 
	   </tr> 

	   <tr> 
			   <td class="dLabel1">Testing Date</td>
			   <td class="entry1"><Input Name="TestingDate" Prototype="MM/DD/YYYY" type="text" value="<%=datTestingDateMMDDYY%>" size="12" maxlength="12"  />&nbsp;
               <a href="javascript:ShowCalendar('document.EditTS.TestingDate', document.EditTS.TestingDate.value);"><img src="calendar.gif"  align="absmiddle" width="18" height="18" border="0" alt="Click Here to Pick up a Date" /> &nbsp;Select a date</a></td>			   
               <td class="drowerror1">&nbsp;</td>   
	   
	   </tr>
	   <tr>
	     <td>Test Status <%=strTestingStatus%></td>
		 <td>

		    <%
			response.write("Status:"&strTestingStatus&":")
			
		    Call SelectFromTable("TestSteps",strTestingStatus)
	        %>
		 </td>
		 <td>&nbsp;</td>
	 
       </tr> 
	   <tr>
	     <td>Test Active</td>
		 <td>

		    <%
		     Call SelectCheckbox("TestIsActive",bitTestIsactive)
	         %>
		 </td>
		 <td><%=bitTestActive%></td>
		 <td>&nbsp;</td>
		 
	   </tr> 
	   <tr> 
	     <td>Last Updated on</td>
		 <td><%=datLastChangeDate%></td> 
		 <td>&nbsp;</td>
		 
	   </tr>
      <tr> 
	     <td>Last Updated by</td>
		 <td><%=strUpdatedBy%></td> 
		 <td>&nbsp;</td>
		 
	   </tr>
       <tr>
	   	 <td>&nbsp;</td>
		 <td>&nbsp;</td>
		 <td>&nbsp;</td>
       </tr>
       <tr>
	   	 <td><input type="reset" name="reset" value="Reset" /></td>
		 <td>&nbsp;</td>
		 <td><input type="submit" name="submit" value="Submit" /></td>
       </tr>
	   
	</form>
	   <tr>
	   	 <td colspan = "3">&nbsp;</td>
       </tr>
	   <tr>
	   	 <td colspan = "3"><a href="DelTestingSession.asp?tid=<%=strTID%>">Delete Test Session <%=strTID%></a></td>
       </tr>

   </table>
        </div> 
      </div>  
	</div>
        
	</body>
</html>
