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

StartTimer 1
bitDebug = 0
'
'  Load Kata Rank Division Table
'

dim KRDTCountByDiv(4)
dim KRDTRankID(4,30)
dim KRDTRankName(4,30)
dim KRDTRankIDRev(4,30)

KRDTCountByDiv(1) = 14
KRDTCountByDiv(2) = 14
KRDTCountByDiv(3) = 11
KRDTCountByDiv(4) = 13

KRDTRankID(1,1)  = 1
KRDTRankID(1,2)  = 2
KRDTRankID(1,3)  = 3
KRDTRankID(1,4)  = 4
KRDTRankID(1,5)  = 5
KRDTRankID(1,6)  = 6
KRDTRankID(1,7)  = 7
KRDTRankID(1,8)  = 8
KRDTRankID(1,9)  = 9
KRDTRankID(1,10) = 10
KRDTRankID(1,11) = 11
KRDTRankID(1,12) = 12
KRDTRankID(1,13) = 13
KRDTRankID(1,14) = 14

KRDTRankID(2,1)  = 1
KRDTRankID(2,2)  = 2
KRDTRankID(2,3)  = 3
KRDTRankID(2,4)  = 4
KRDTRankID(2,5)  = 5
KRDTRankID(2,6)  = 6
KRDTRankID(2,7)  = 7
KRDTRankID(2,8)  = 8
KRDTRankID(2,9)  = 9
KRDTRankID(2,10) = 10
KRDTRankID(2,11) = 11
KRDTRankID(2,12) = 12
KRDTRankID(2,13) = 13
KRDTRankID(2,14) = 14

KRDTRankID(3,1)  = 1
KRDTRankID(3,2)  = 2
KRDTRankID(3,3)  = 3
KRDTRankID(3,4)  = 5
KRDTRankID(3,5)  = 7
KRDTRankID(3,6)  = 9
KRDTRankID(3,7)  = 11
KRDTRankID(3,8)  = 12
KRDTRankID(3,9)  = 13
KRDTRankID(3,10) = 14
KRDTRankID(3,11) = 15

KRDTRankID(4,1)  = 1
KRDTRankID(4,2)  = 2
KRDTRankID(4,3)  = 5
KRDTRankID(4,4)  = 11
KRDTRankID(4,5)  = 12
KRDTRankID(4,6)  = 13
KRDTRankID(4,7)  = 15
KRDTRankID(4,8)  = 16
KRDTRankID(4,9)  = 17
KRDTRankID(4,10) = 18
KRDTRankID(4,11) = 19
KRDTRankID(4,12) = 23
KRDTRankID(4,13) = 24


KRDTRankName(1,1)  = "White Belt II"
KRDTRankName(1,2)  = "Go Kyu"
KRDTRankName(1,3)  = "Orange Belt"
KRDTRankName(1,4)  = "Orange Belt 2"
KRDTRankName(1,5)  = "Yon Kyu"
KRDTRankName(1,6)  = "Green Belt II"
KRDTRankName(1,7)  = "Blue Belt"
KRDTRankName(1,8)  = "Blue Belt II"
KRDTRankName(1,9)  = "Purple Belt"
KRDTRankName(1,10) = "Purple Belt II"
KRDTRankName(1,11) = "San Kyu"
KRDTRankName(1,12) = "Ni Kyu"
KRDTRankName(1,13) = "Ikkyu"
KRDTRankName(1,14) = "Junior BB"
KRDTRankName(2,1)  = "White Belt II"
KRDTRankName(2,2)  = "Go Kyu"
KRDTRankName(2,3)  = "Orange Belt"
KRDTRankName(2,4)  = "Orange Belt 2"
KRDTRankName(2,5)  = "Yon Kyu"
KRDTRankName(2,6)  = "Green Belt II"
KRDTRankName(2,7)  = "Blue Belt"
KRDTRankName(2,8)  = "Blue Belt II"
KRDTRankName(2,9)  = "Purple Belt"
KRDTRankName(2,10) = "Purple Belt II"
KRDTRankName(2,11) = "San Kyu"
KRDTRankName(2,12) = "Ni Kyu"
KRDTRankName(2,13) = "Ikkyu"
KRDTRankName(2,14) = "Junior BB"
KRDTRankName(3,1)  = "White Belt II"
KRDTRankName(3,2)  = "Go Kyu"
KRDTRankName(3,3)  = "Orange Belt"
KRDTRankName(3,4)  = "Yon Kyu"
KRDTRankName(3,5)  = "Blue Belt"
KRDTRankName(3,6)  = "Purple Belt"
KRDTRankName(3,7)  = "San Kyu"
KRDTRankName(3,8)  = "Ni Kyu"
KRDTRankName(3,9)  = "Ikkyu"
KRDTRankName(3,10) = "Junior BB"
KRDTRankName(3,11) = "Shodan"
KRDTRankName(4,1)  = "White Belt II"
KRDTRankName(4,2)  = "Go Kyu"
KRDTRankName(4,3)  = "Yon Kyu"
KRDTRankName(4,4)  = "San Kyu"
KRDTRankName(4,5)  = "Ni Kyu"
KRDTRankName(4,6)  = "Ikkyu"
KRDTRankName(4,7)  = "Shodan"
KRDTRankName(4,8)  = "Nidan"
KRDTRankName(4,9)  = "Sandan"
KRDTRankName(4,10) = "Yo dan"
KRDTRankName(4,11) = "Go dan"
' KRDTRankName(4,13) = "Rokudan"
' KRDTRankName(4,14) = "Nanadan"
' KRDTRankName(4,15) = "Hachidan"
KRDTRankName(4,12) = "Kudan"
KRDTRankName(4,13) = "Judan"





'Reverse lookup dump
KRDTRankIDRev( 1,1) = 1
KRDTRankIDRev( 1,2) = 2
KRDTRankIDRev( 1,3) = 3
KRDTRankIDRev( 1,4) = 4
KRDTRankIDRev( 1,5) = 5
KRDTRankIDRev( 1,6) = 6
KRDTRankIDRev( 1,7) = 7
KRDTRankIDRev( 1,8) = 8
KRDTRankIDRev( 1,9) = 9
KRDTRankIDRev( 1,10) = 10
KRDTRankIDRev( 1,11) = 11
KRDTRankIDRev( 1,12) = 12
KRDTRankIDRev( 1,13) = 13
KRDTRankIDRev( 1,14) = 14
KRDTRankIDRev( 1,15) = 0
KRDTRankIDRev( 1,16) = 0
KRDTRankIDRev( 1,17) = 0
KRDTRankIDRev( 1,18) = 0 
KRDTRankIDRev( 1,19) = 0
KRDTRankIDRev( 1,20) = 0
KRDTRankIDRev( 1,21) = 0
KRDTRankIDRev( 1,22) = 0
KRDTRankIDRev( 1,23) = 0
KRDTRankIDRev( 1,24) = 0
KRDTRankIDRev( 1,25) = 0
KRDTRankIDRev( 1,26) = 0
KRDTRankIDRev( 1,27) = 0
KRDTRankIDRev( 1,28) = 0
KRDTRankIDRev( 1,29) = 0
KRDTRankIDRev( 1,30) = 0

KRDTRankIDRev( 2,1) = 1
KRDTRankIDRev( 2,2) = 2
KRDTRankIDRev( 2,3) = 3
KRDTRankIDRev( 2,4) = 4
KRDTRankIDRev( 2,5) = 5
KRDTRankIDRev( 2,6) = 6
KRDTRankIDRev( 2,7) = 7
KRDTRankIDRev( 2,8) = 8
KRDTRankIDRev( 2,9) = 9
KRDTRankIDRev( 2,10) = 10
KRDTRankIDRev( 2,11) = 11
KRDTRankIDRev( 2,12) = 12
KRDTRankIDRev( 2,13) = 13
KRDTRankIDRev( 2,14) = 14
KRDTRankIDRev( 2,15) = 0
KRDTRankIDRev( 2,16) = 0 
KRDTRankIDRev( 2,17) = 0
KRDTRankIDRev( 2,18) = 0
KRDTRankIDRev( 2,19) = 0
KRDTRankIDRev( 2,20) = 0
KRDTRankIDRev( 2,21) = 0
KRDTRankIDRev( 2,22) = 0
KRDTRankIDRev( 2,23) = 0
KRDTRankIDRev( 2,24) = 0
KRDTRankIDRev( 2,25) = 0
KRDTRankIDRev( 2,26) = 0
KRDTRankIDRev( 2,27) = 0
KRDTRankIDRev( 2,28) = 0
KRDTRankIDRev( 2,29) = 0
KRDTRankIDRev( 2,30) = 0

KRDTRankIDRev( 3,1) = 1
KRDTRankIDRev( 3,2) = 2
KRDTRankIDRev( 3,3) = 3
KRDTRankIDRev( 3,4) = 0
KRDTRankIDRev( 3,5) = 4
KRDTRankIDRev( 3,6) = 0
KRDTRankIDRev( 3,7) = 5
KRDTRankIDRev( 3,8) = 0
KRDTRankIDRev( 3,9) = 6
KRDTRankIDRev( 3,10) = 0
KRDTRankIDRev( 3,11) = 7
KRDTRankIDRev( 3,12) = 8
KRDTRankIDRev( 3,13) = 9
KRDTRankIDRev( 3,14) = 10
KRDTRankIDRev( 3,15) = 11
KRDTRankIDRev( 3,16) = 0
KRDTRankIDRev( 3,17) = 0
KRDTRankIDRev( 3,18) = 0
KRDTRankIDRev( 3,19) = 0
KRDTRankIDRev( 3,20) = 0
KRDTRankIDRev( 3,21) = 0
KRDTRankIDRev( 3,22) = 0
KRDTRankIDRev( 3,23) = 0
KRDTRankIDRev( 3,24) = 0
KRDTRankIDRev( 3,25) = 0
KRDTRankIDRev( 3,26) = 0
KRDTRankIDRev( 3,27) = 0
KRDTRankIDRev( 3,28) = 0
KRDTRankIDRev( 3,29) = 0
KRDTRankIDRev( 3,30) = 0

KRDTRankIDRev( 4,1) = 1
KRDTRankIDRev( 4,2) = 2
KRDTRankIDRev( 4,3) = 0
KRDTRankIDRev( 4,4) = 0
KRDTRankIDRev( 4,5) = 3
KRDTRankIDRev( 4,6) = 0
KRDTRankIDRev( 4,7) = 0
KRDTRankIDRev( 4,8) = 0
KRDTRankIDRev( 4,9) = 0
KRDTRankIDRev( 4,10) = 0
KRDTRankIDRev( 4,11) = 4
KRDTRankIDRev( 4,12) = 5
KRDTRankIDRev( 4,13) = 6
KRDTRankIDRev( 4,14) = 0
KRDTRankIDRev( 4,15) = 7
KRDTRankIDRev( 4,16) = 8
KRDTRankIDRev( 4,17) = 9
KRDTRankIDRev( 4,18) = 10
KRDTRankIDRev( 4,19) = 11
KRDTRankIDRev( 4,20) = 0
KRDTRankIDRev( 4,21) = 0
KRDTRankIDRev( 4,22) = 0
KRDTRankIDRev( 4,23) = 12
KRDTRankIDRev( 4,24) = 13
KRDTRankIDRev( 4,25) = 0
KRDTRankIDRev( 4,26) = 0
KRDTRankIDRev( 4,27) = 0
KRDTRankIDRev( 4,28) = 0
KRDTRankIDRev( 4,29) = 0
KRDTRankIDRev( 4,30) = 0


oRSr.ActiveConnection = oConn 
            

                   strusername = Session("Username")
                   strSQLA = "SELECT DojoCd, CanChangeDojo, SelActive, SelInactive, TestingID  FROM dms_pref WHERE username='"&strUserName&"';"
			       oRSA.Open strSQLA
                   strDojoCd        = trim(orsA("DojoCd"))
                   strSelActive     = orsA("SelActive")
                   strSelInactive   = orsA("SelInActive")
                   bitCanChangeDojo = orsA("CanChangeDojo")
                   intTestingID     = orsA("TestingID")
                   orsA.Close
                   set orsA = Nothing
                   Set oRSA = Server.CreateObject("ADODB.Recordset") 
                   oRSA.ActiveConnection = oConn 
                   
                   strSQLB = "SELECT * FROM dms_TestSession WHERE TestingID = '"&intTestingID&"';"
                    'response.write("strSQLB=>"&strsqlB&"<</br>")
                   oRSB.Open strSQLB
                   strSponsoringSchoolCd   = orsB("SponsoringSchoolCd")
                   strSponsoringSchoolName = GetDojoName(strSponsoringSchoolCd)
                   strTestingSchools       = orsB("TestingSchools")
                   strTestingSchoolsWhere = ""
                   intSchoolsCnt = 0
                   for inti = 1 to len(strTestingSchools)
                      if mid(strTestingSchools,inti,1) = "1" then
                         intSchoolsCnt = intSchoolsCnt + 1
                         if Len(strTestingSchoolsWhere) > 0 then strTestingSchoolsWhere = strTestingSchoolsWhere & " OR " end if
                         if inti = 1 then strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'npk' " end if
                         if inti = 2 then strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'bk ' " end if
                         if inti = 3 then strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'kin' " end if
                         if inti = 4 then strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'pv ' " end if
                         if inti = 5 then strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'ef ' " end if
                      end if
                   next
                   'response.write("strTestingSchools=>"&strTestingSchools&"<; strTestingSchoolsWhere=>"&strTestingSchoolsWhere&"<<br>")
                   strTestLocationCd       = orsB("TestingLocationCd")
                   strTestingLocationName  = GetCodeValueText("TestLocation", strTestLocationCd)
                   datTestingDate          = orsB("TestingDate")
                   strTestingStatus        = orsB("TestingStatus")
                   orsB.Close
                   set orsB = Nothing
                   Set oRSB = Server.CreateObject("ADODB.Recordset") 
                   oRSB.ActiveConnection = oConn 
                  
                 
                   
                   
			   strSQL = "Select TestingID, dms_TestPull.StudentID , LastTestingDate,  RankID, NextRankID, dms_TestPull.Division, dms_TestPull.Anchorday, AnchorDayOrder, StatusID, AddType, LName, FName, MName FROM dms_TestPull " 
			   strSQL = strSQL & "INNER JOIN dms_Student ON dms_Student.StudentID = dms_TestPull.StudentID "
			   strSQL = strSQL & "WHERE TestingID = '"&intTestingID&"' "
               strSQL = strSQL & "ORDER BY Division, AnchorDayOrder, LName, FName, MName ;"
			   'response.write("strSQL=>"&strSQL&"<<br>")
 %>			   
			   <table>
                  <tr>
                     <td colspan="12"><%=strSponsoringSchoolname%> Testing ID is <%=intTestingID%></td>
                  </tr>
			      <tr>
	          <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			        <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">Last</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>                  
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
	  	  	      </tr>    
                   
			      <tr>
	          <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			        <td class="dtabh">Student Name</td>
                    <td class="dtabh">Last Rank</td>
                    <td class="dtabh">Test Date</td>
                    <td class="dtabh" >Status</td>
                    <td class="dtabh" >&nbsp;</td>
                    <td class="dtabh" >Next Rank</td>
                    <td class="dtabh" >&nbsp;</td>
                    <td class="dtabh" >&nbsp;</td>                  
                    <td class="dtabh">Div</td>
                    <td class="dtabh"><i class="fa fa-anchor"></td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
	  	  	      </tr>    
                  
			<%
			oRS.Open strSQL
			intCnt = 0
		    do while not ors.eof
		    intCnt = intCnt + 1
			   strLname            = ors("LName")
			   strfname            = ors("FName")
			   strMName            = ors("MName")
			   strStudentName      = left(Trim(strFName)&" "&Trim(strLName), 32)
			   intStudentID        = ors("StudentID")
			   strDivision         = ors("Division")
			   strAnchorDay        = ors("AnchorDay")
			   datLastTested       = ors("LastTestingDate")
			   strRankCd           = ors("RankId")
			   intRankID           = CInt(strRankCd)
			  ' strRankName         = GetRankName(strRankCd)&"     "
			   if intRankID = 0 then
			      strRankName = "White Belt     "
			   else    
			      strRankName         = KRDTRankName(intD, KRDTRankIDRev(intD,intRankID ))&"     "
			   end if   
			   intRCurr            = intRankID
			   ' *************************************************************************************************
			   strNextRankName     = "***"
			   intD                = cint(strDivision)
	 		   intR                = Cint(ors("NextRankId"))
	 		   intNextRankID       = intR
	 		   strNextRankName     = KRDTRankName(intD, KRDTRankIDRev(intD,intR))&"     "
			   'datLastTested       = GetLastTestingDate(intStudentID)
			   strStatusID         = ors("StatusID")
			   if bitDebug then response.write("<br>"&strLname&","&strFName&"-"&intStudentID&"<<br>") end if
			   
               bitMatch = 0
               intii = 0
               if intRankId = intNextRankID  then 
			      intNextRankIDStart   = intRankID
			      strNextRankNameStart = strNextRankName
			   else 
			      intNextRankIDStart   = KRDTRankIDRev(intD,intRCurr)+1
			      strNextRankNameStart = KRDTRankName(intD,intNextRankIDStart)
			   end if   
			   if bitDebug = 1 then response.write(intStudentID&" "&strLName&","&strFName&" intNextRankID=>"&intNextRankID&"<; <br>") end if
			%>
			 <tr>
	     <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			      <form form="form<%=intStudentID%>" id="updatepullline<%=intStudentID%>" action="updatepullline.asp?ts=<%=intTestingID%>" method="post">

			      <td class="dtab" id="SN<%=intStudentID%>"><%=strStudentName%></td>
                  <td class="dtab">&nbsp;<%=strRankName%>&nbsp;</td>
                  <td class="dtab">&nbsp;<%=datLastTested%>&nbsp;</td>
                  <% strJS = "onchange="&chr(34)&"document.getElementById('updatepullline"&intStudentID&"').submit();"&chr(34)%>
                  <td class="dtab"><% Call SelectFromTableIndJS("TestingStatus",strStatusID,intStudentID,strJS)%></td>  
                  <td class="dtab"><input name="recommend" type="button" value = "R" onclick="selectFunction('TestingStatus<%=intStudentID%>','4');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                   <input name="excuse"    type="button" value = "E" onclick="selectFunction('TestingStatus<%=intStudentID%>','3');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                   <input name="Held"      type="button" value = "H" onclick="selectFunction('TestingStatus<%=intStudentID%>','2');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                   <input name="promote"   type="button" value = "P" onclick="selectFunction('TestingStatus<%=intStudentID%>','1');document.getElementById('updatepullline<%=intStudentID%>').submit()"> &nbsp; 
                  </td>
                                  
                  <td class="dtab"> &nbsp;
                     <select name = "NextRankID<%=intStudentID%>" id="NextRankID<%=intStudentID%>" onchange="document.getElementById('updatepullline<%=intStudentID%>').submit();">
                  <%
                        intD = CInt(strDivision)
                        if bitDebug = 1 then response.write("^^^^^^^^^^^ "&intNextRankIDStart&" KRDTCountbyDiv("&intd&")="&KRDTCountbyDiv(intD)&"<br>") end if
                        if intRankId = intNextRankID then 
                        %>
                        <option value="<%= intNextRankIDStart%>" SELECTED ><%=strRankName%></option>
                        <%
                        else
                           for inti = intNextRankIDStart  to KRDTCountbyDiv(intD)
                              if bitDebug = 1 then response.write("##  KRDTRankIDRev("&intD&","&inti&")=>"& KRDTRankIDRev(intD,inti)&"> KRDTRankID("&intD&","&inti&")="&KRDTRankID(intD,inti)&" nextRankID="&Cint(intNextRankId)&"<br>") end if
                              if  KRDTRankID(intD,inti) = CInt(intNextRankID) then
                                  strSELECTED = "SELECTED"
                              else
                                  strSELECTED = ""
                              end if        
                  %>   
                              <option value="<%= inti%>" <%=strSELECTED%> ><%=KRDTRankName(intD,inti)%></option>
                  <%
                        next
                       end if 
                  %>
                     </select>
                  </td>
                  <td class="dtab"><input name="Plus"  type="button" value = "+" onclick="selectFunctionInc('NextRankID<%=intStudentID%>',<%=KRDTCountbyDiv(intD)%>);document.getElementById('updatepullline<%=intStudentID%>').submit()"></td>
                  <td class="dtab"><input name="Minus" type="button" value = "-" onclick="selectFunctionDec('NextRankID<%=intStudentID%>',<%=intNextRankIdStart%>);document.getElementById('updatepullline<%=intStudentID%>').submit()"></td>                  
                  <td class="dtab" align="center" ><%=strDivision%></td>
                  <td class="dtab" align="center" ><%=strAnchorDay%></td>
                  <td class="dtab"><input name="Div<%=intStudentID%>" type="hidden" value="<%=strDivision%>" >
                                   <input name="SID<%=intStudentID%>" type="hidden" value="<%=intStudentID%>" >
                                   <input name="Submit<%=intStudentID%>" type="submit" id="submit-button<%=intStudentID%>"  style="display: none" onclick="window.open('updatepulline.asp.asp?ts=<%=intTestingID%>', '_parent', 'width=1000,height=400,menubar=yes'); return false; ">
                                  
                                   </td>
                  <td class="dtab">&nbsp;</td>
			 </tr> 
		</form>  

            <%
			ors.Movenext
			loop
			Elapsed = StopTimer(1)
            Response.Write "Elapsed time was: " & Elapsed 
			%>
            <tr>
               <td class="dtabh">&nbsp;</td>
               <td colspan="3" class="dtabh">&nbsp;<%=intCnt%> Students</td>
               <td class="dtabh">&nbsp;</td>
            </tr>
            <tr>
               <td colspan="5" class="dtabnone">&nbsp;</td>
            </tr>
            <tr>
               <td class="dtabnone">&nbsp;</td>
               <td colspan="5"  align="left" valign="middle" class="dtabnone"></td>
            
           </tr> 
		 </TABLE>
         <br/><br/>
		</div>	
      </div>
	</div>
</body>
</html>
