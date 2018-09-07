<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Edit Pulled Data</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
<script  type="text/javascript">

var selectFunction = function(bname, bvalue) {
    document.getElementById(bname).value = bvalue;
    };
    
var selectFunctionInc = function(bname, maxvalue) {
	currrank = document.getElementById(bname).value;
	currrank = parseInt(currrank);
	maxvalue = parseInt(maxvalue);
	console.log(currrank);
	console.log(maxvalue);
	if (currrank < maxvalue) {
	   currrank = currrank + 1;
	   currrank = currrank + "";
    }     
	document.getElementById(bname).value = currrank;
};  
var selectFunctionDec = function(bname, minvalue) {
	currrank = document.getElementById(bname).value;
	currrank = parseInt(currrank);
	minvalue = parseInt(minvalue);
	console.log(currrank);
	console.log(minvalue);
	if (currrank > minvalue) {
	   currrank = currrank - 1;
	   currrank = currrank + "";
    }     
	document.getElementById(bname).value = currrank;
};   

function showHint(str) {
    if (str.length == 0) { 
        document.getElementById("txtHint").innerHTML = "";
        return;
    } else {
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (this.readyState == 4 && this.status == 200) {
                document.getElementById("txtHint").innerHTML = this.responseText;
            }
        };
        xmlhttp.open("GET", "gethint1.asp?q=" + str, true);
        xmlhttp.send();
    }
}
</script>   

</head>

<body>
<div id="main">
	  <div id="contents">
	  <div id="OneCol">
      <br/>
<%
Dim StopWatch(19) 
sub StartTimer(x)
'::::::::::::::::::::::::::::::::::::::::::::::::::::
':::     Here we start the timer                  :::
'::::::::::::::::::::::::::::::::::::::::::::::::::::
   StopWatch(x) = timer
end sub


function StopTimer(x)
'::::::::::::::::::::::::::::::::::::::::::::::::::::
':::     Here we 'stop' the timer and calculate   :::
':::     the elapsed time (allowing for the       :::
':::     midnight wrap-around. NOTE: These        :::
':::     routines should not be used to time      :::
':::     very long events (>1 hour) or very       :::
':::     short events (< 100 milliseconds)        :::
'::::::::::::::::::::::::::::::::::::::::::::::::::::
   EndTime = Timer

   'Watch for the midnight wraparound...
   if EndTime < StopWatch(x) then
     EndTime = EndTime + (86400)
   end if

   StopTimer = EndTime - StopWatch(x)
end function
'response.write("<script> alert('This may take a minute or so, please wait...!');</script>")
StartTimer 1
bitDebug = 0

'
'  Load Kata Rank Division Table
'


intTID    = Request.QueryString("TID")
strDojoCd = Request.QueryString("DC")
strDojoCdSave = Trim(strDojoCd)
strDojoCd = left(strDojoCd&"   ",3)
intFR     = Request.QueryString("FR")
intTR     = Request.QueryString("TR")
strSD     = Request.QueryString("SD")
ints      = CInt(Request.QueryString("s"))

'response.write("TID=>"&intTID&"<, DC=>"&strDC&"<, FR=>"&intFR&"<, TR=>"&intTR&"<, SD=>"&strSD&"<, s=>"&ints&"< <br>")
 
Dim RankNaMES(28)

strSQL = "SELECT CodeName, CodeValue, CodeValueText, ValueOrder FROM dms_Codes WHERE CodeName = 'RankCd' AND ValueIsActive = 1 ORDER BY ValueOrder;"
orsB.Open strSQL
do while NOT orsB.EOF
   intRN    = orsB("CodeValue")
   intRName = orsB("CodeValueText")
   RankNames(intRN) = intRName
   orsB.MoveNext
loop
                   orsB.Close
                   set orsB = Nothing
                   Set oRSB = Server.CreateObject("ADODB.Recordset") 
                   oRSB.ActiveConnection = oConn 

         
strSDWhere = ""
strSDstr   = "Divisions: "
                   for inti = 1 to len(strSD)
                      if mid(strSD,inti,1) = "1" then
                         intSchoolsCnt = intSchoolsCnt + 1
                         if Len(strSDWhere) > 0 then strSDWhere = strSDWhere & " OR " end if
                         if Len(strSDStr) > 0   then strSDStr = strSDstr & ", " end if
                         if inti = 1 then 
                            strSDWhere = strSDWhere & "  dms_TestPull.Division = '1' " 
                            strSDStr = strSDStr & "1 "
                         end if
                         if inti = 2 then 
                            strSDWhere = strSDWhere & "  dms_TestPull.Division = '2' " 
                            strSDStr = strSDStr & "2 "
                         end if
                         if inti = 3 then
                            strSDWhere = strSDWhere & "  dms_TestPull.Division = '3' "  
                            strSDStr = strSDStr & "3  "
                         end if
                         if inti = 4 then 
                            strSDWhere = strSDWhere & "  dms_TestPull.Division = '4' "
                            strSDStr = strSDStr & "4  "
                         end if
                      end if
                   next
                   strSQLTS = "SELECT * FROM dms_TestSession WHERE TestingID = "&intTID&";"
                   orsB.open strSQLTS
                   if orsB.EOF then
                      response.write("Test Session "&intTID&", not found!<br>")
                      response.end
                   end if
                   'response.write("strSD=>"&strSD&"<; strSDWhere=>"&strSDWhere&"<<br>") 
                   strSponsor              = orsB("SponsoringSchoolCd")
                   strTestLocationCd       = orsB("TestingLocationCd")
                   strTestingLocationName  = GetCodeValueText("TestLocation", strTestLocationCd)
                   datTestingDate          = orsB("TestingDate")
                   strTestingStatus        = orsB("TestingStatus")
                   strTestingSchools       = orsB("TestingSchools")
                   strTestingSchoolsStr   = ""
                   intSchoolsCnt = 0
                   for inti = 1 to len(strTestingSchools)
                      if mid(strTestingSchools,inti,1) = "1" then
                         intSchoolsCnt = intSchoolsCnt + 1
                         if Len(strTestingSchoolsStr) > 0 then strTestingSchoolsStr = strTestingSchoolsstr & ", " end if
                         if inti = 1 then 
                            strTestingSchoolsStr = strTestingSchoolsStr & "New Paltz "
                         end if
                         if inti = 2 then 
                            strTestingSchoolsStr = strTestingSchoolsStr & "Brooklyn "
                         end if
                         if inti = 3 then
                            strTestingSchoolsStr = strTestingSchoolsStr & "Kinnelon "
                         end if
                         if inti = 4 then 
                            strTestingSchoolsStr = strTestingSchoolsStr & "Pleasant Valley "
                         end if
                         if inti = 5 then 
                            strTestingSchoolsStr = strTestingSchoolsStr & "East Fishkill "
                         end if
                      end if
                   next
                   'response.write("strTestingSchools=>"&strTestingSchools&"<; strTestingSchoolsWhere=>"&strTestingSchoolsWhere&"<<br>")
                  
                   
                   
                   orsB.Close
                   set orsB = Nothing
                   Set oRSB = Server.CreateObject("ADODB.Recordset") 
                   oRSB.ActiveConnection = oConn 
                   strSortOrderCd = intS
                   
                   Select Case strSortOrderCd
                          Case "0"
                             strSortOrder = "ORDER BY dms_TestPull.Division, AnchorDayOrder, LName, FName, MName"
                          Case "1"
                             strSortOrder = "ORDER BY LName, FName, MName, dms_TestPull.Division, AnchorDayOrder  "
                          Case "2"
                             strSortOrder = "ORDER BY RankID, LName, FName, MName, dms_TestPull.Division, AnchorDayOrder  "
                          Case "3"
                             strSortOrder = "ORDER BY LastTestingDate, LName, FName, MName, dms_TestPull.Division, AnchorDayOrder  "
                          Case "4"
                             strSortOrder = "ORDER BY StatusID, LName, FName, MName, dms_TestPull.Division, AnchorDayOrder  "
                          Case "5"
                             strSortOrder = "ORDER BY NextRankID, LName, FName, MName, dms_TestPull.Division, AnchorDayOrder  "
                          Case "6"
                             strSortOrder = "ORDER BY dms_TestPull.Division,  LName, FName, MName,  AnchorDayOrder  "
                          Case "7"
                             strSortOrder = "ORDER BY AnchorDayOrder,StatusID, LName, FName, MName, dms_TestPull.Division "
                          Case "8"
                             strSortOrder = "ORDER BY dms_TestPull.DojoCd, dms_TestPull.Division, AnchorDayOrder, LName, FName, MName"
                            
                          Case else
                             strSortOrder = "ORDER BY dms_TestPull.Division, AnchorDayOrder, LName, FName, MName"
End Select  
                 
                   
                   
			   strSQL = "Select TestingID, dms_TestPull.StudentID , dms_TestPull.DojoCd AS TPDojoCd,  dms_TestPull.LastTestingDate,  RankID, NextRankID, dms_TestPull.Division AS Div, dms_TestPull.Anchorday, AnchorDayOrder, StatusID, AddType, LName, FName, MName FROM dms_TestPull " 
			   strSQL = strSQL & "INNER JOIN dms_Student ON dms_Student.StudentID = dms_TestPull.StudentID "
			   strSQL = strSQL & "WHERE NOT AddType= 'O' AND TestingID = '"&intTID&"'  AND NextRankID >= "&intFR&"  AND NextRankID <= "&intTR&" AND StatusID = 4 AND dms_TestPull.DojoCd = '"&strDojoCdSave&"' AND ("&strSDWhere&") "
               strSQL = strSQL & strSortorder& " ;"
			 '  response.write("strSQL=>"&strSQL&"<<br>") 
 %>
 <h1> Display Group Information</h1>

 <h2>&nbsp;Sponsor: <%=strSponsor%></h2>		
  	   
			   <table>
                  <tr>
                     <td colspan="12"> <b>Participating Schools:</b> <%= strTestingSchoolsStr %></td>
                  </tr>
			   
                  <tr>
                     <td colspan="12"> <b>Test Date:</b>  <%=datTestingDate%>&nbsp;&nbsp; <b>Test Location:</b><%=strTestingLocationName%>&nbsp;&nbsp;<b> Testing ID:</b><%=intTID%>&nbsp;&nbsp;&nbsp;  <a href="Testing.asp"><input type="button" Name="Goback" Value="Close" onclick='self.close(); window.onunload = refreshParent;'></a>
</td>
                  </tr>
			      <tr>
			        <td class="dtabh">&nbsp;</td>
	          		<td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>                  
	  	  	      </tr>    
 <%
 strArguments = "&TID="&intTID&"&DC="&Trim(strDojoCdSave)&"&FR="&intFR&"&TR="&intTR&"&SD="&strSD

 %>                  
			      <tr>
			        <td class="dtabh"><a href="DisplayGroup.asp?s=8<%=strArguments%>">Dojo</a></td>
       	            <td class="dtabh"><a href="DisplayGroup.asp?s=1<%=strArguments%>">Student Name</a></td>
                    <td class="dtabh"><a href="DisplayGroup.asp?s=2<%=strArguments%>">Last Rank</a></td>
                    <td class="dtabh"><a href="DisplayGroup.asp?s=3<%=strArguments%>">Test Date</a></td>
                    <td class="dtabh"><a href="DisplayGroup.asp?s=4<%=strArguments%>">Status</a></td>
                    <td class="dtabh"><a href="DisplayGroup.asp?s=5<%=strArguments%>">Next Rank</a></td>
                    <td class="dtabh"><a href="DisplayGroup.asp?s=6<%=strArguments%>">Div</a></td>
                    <td class="dtabh"><a href="DisplayGroup.asp?s=7<%=strArguments%>"><i class="fa fa-anchor"></a></td>
	  	  	      </tr>    
                  
			<%
			oRS.Open strSQL
			intCnt = 0
			if ors.eof then
			   response.write("EOF EARLY!!!<br>")
			   response.end
			end if   
		    do while not ors.eof
		       intCnt = intCnt + 1
			   strLname            = ors("LName")
			   strfname            = ors("FName")
			   strMName            = ors("MName")
			   strStudentName      = left(Trim(strFName)&" "&Trim(strLName), 32)
			   intStudentID        = ors("StudentID")
			   strDojoCd           = ors("TPDojoCd")
			   strDivision         = ors("Div")
			   strAnchorDay        = ors("AnchorDay")
			   intStatusID         = ors("StatusID")
			   datLastTested       = ors("LastTestingDate")
			   strRankCd           = ors("RankId")
			   intRankID           = CInt(strRankCd)
			   strRankName         = RankNames(intRankID)
			   
	 		   strNRankCd          = ors("NextRankId") 
			   intNRankID          = CInt(strNRankCd)
			   strNRankName        = RankNames(intNRankID)
	 		   
	 		   
			 '  strNRankName        = GetRankName(strNRankCd)
			   
			   if bitDebug = 1 then response.write(intStudentID&" "&strLName&","&strFName&" intNextRankID=>"&intNextRankID&"<; <br>") end if
			%>		
			 <tr>
			      <td class="dtab"><%=strDojoCd%></td>
			      <td class="dtab"><%=strStudentName%></td>
                  <td class="dtab">&nbsp;<%=strRankName%>&nbsp;</td>
                  <td class="dtab">&nbsp;<%=datLastTested%>&nbsp;</td>
                  <td class="dtab">&nbsp;<%=intStatusID%></td>                  
                  <td class="dtab"><%=strNRankName%></td>
                  <td class="dtab" align="center" ><%=strDivision%></td>
                  <td class="dtab" align="center" ><%=strAnchorDay%></td>
                 
			 </tr> 
		 

            <%
			ors.Movenext
			loop
			
			%>
            <tr>
               <td class="dtabh">&nbsp;</td>
               <td colspan="3" class="dtabh">&nbsp;<%=intCnt%> Students</td>
               <td class="dtabh">&nbsp;</td>
               <td class="dtabh">&nbsp;</td>
               <td class="dtabh">&nbsp;</td> 
               <td class="dtabh">&nbsp;</td>
            </tr>
		 </TABLE>
		 
         <br/><br/>
		</div>	
      </div>
      <%
      Elapsed = StopTimer(1)
      Response.Write "Elapsed time was: " & Elapsed 
      %>
	</div>
</body>
</html>
