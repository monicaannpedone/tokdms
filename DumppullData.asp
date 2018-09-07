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

bitDebug = 1
for inti = 1 to 4
    if bitDebug = 1 then  response.write("inti = "&inti&"; Count for Div "&inti&" is "& KRDTCountbyDiv(inti)&"<br>") end if
    for intJ = 1 to KRDTCountbyDiv(inti)
        if bitDebug = 1 then response.write("intj = "&intj&"; Rank ID = "& KRDTRankID(inti,intj)&"; "&KRDTRankName(inti,intj)&"<br>") end if
    next
     if bitDebug = 1 then response.write("<br>") end if
next
 if bitDebug = 1 then response.write("<br><br>Reverse lookup dump <br>") end if
oRSr.ActiveConnection = oConn 
for intj = 1 to 4
	for inti = 1 to 30
	    if bitDebug = 1 then response.write("Rev Lookup Div "&intj&","&inti&" -- "&KRDTRankIDRev(intj, inti)&"<br>") end if
	next 
	response.write("<br>")

Next  
            

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
                  
                 
                   
                   
			   strSQL = "Select TestingID, dms_TestPull.StudentID ,dms_TestPull.DojoCd,   RankID, NextRankID, dms_TestPull.Division, dms_TestPull.Anchorday, AnchorDayOrder, StatusID, AddType, LName, FName, MName FROM dms_TestPull " 
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
                    <td class="dtabh">Dojo</td>			        
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
			   strRankCd           = ors("RankId")
			   strDojoCd           = ors("DojoCd")
			   intRankID           = CInt(strRankCd)
			   strRankName         = GetRankName(strRankCd)&"     "
			   ' *************************************************************************************************
			   strNextRankName     = "***"
			   intD = cint(strDivision)
	 		   intR                = Cint(ors("NextRankId"))
	 		   intNextRankID       = Cint(KRDTRankIDRev(intD,intR))
	 		   strNextRankName     = KRDTRankName(intD, intNextRankID)&"     "
			   datLastTested       = GetLastTestingDate(intStudentID)
			   strStatusID         = ors("StatusID")
			  ' response.write("<br>"&strLname&","&strFName&"-"&intStudentID&"<<br>")
			   
			   if bitDebug = 1 then response.write(intStudentID&" "&strLName&","&strFName&" intNextRankID=>"&intNextRankID&"<; <br>") end if
			%>
			 <tr>
	     <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			      <form form="form<%=intStudentID%>" id="updatepullline<%=intStudentID%>" action="updatepullline.asp?ts=<%=intTestingID%>" method="post">
                  <td class="dtab"> &nbsp;<%=strDojoCd%></td>
			      <td class="dtab" id="SN<%=intStudentID%>"><%=strStudentName%></td>
                  <td class="dtab">&nbsp;<%=strRankName%>&nbsp;</td>
                  <td class="dtab">&nbsp;<%=datLastTested%>&nbsp;</td>
                  <td class="dtab"><% Call SelectFromTableInd("TestingStatus",strStatusID,intStudentID)%></td>  
                  <td class="dtab">&nbsp  </td>
                                  
                  <td class="dtab"> <%=intNextRankID%>::<%=strNextRankName%>
                    
                  </td>
                  <td class="dtab">&nbsp;</td>
                  <td class="dtab">&nbsp;</td>                  
                  <td class="dtab" align="center" ><%=strDivision%></td>
                  <td class="dtab" align="center" ><%=strAnchorDay%></td>
                  <td class="dtab">&nbsp;</td>
                  <td class="dtab">&nbsp;</td>
			 </tr> 
		</form>  

            <%
			ors.Movenext
			loop
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
