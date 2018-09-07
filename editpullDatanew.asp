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
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>	
<script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.6.4/angular.min.js"></script>
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
               strMsg = Request.QueryString("Msg")
               strSelDojoCd = Request.Querystring("SelDojoCd")
              ' response.write("Request.Querystring(SelDojoCd)=>"&Request.Querystring("SelDojoCd")&"<<br>")
               strselDojoCd = Left(strSelDojoCd&"   ",3)
               strSelDiv1   = Request.Querystring("SelDiv1")
               strSelDiv2   = Request.Querystring("SelDiv2")
               strSelDiv3   = Request.Querystring("SelDiv3")
               strSelDiv4   = Request.Querystring("SelDiv4")
               
               intSelGF     = CInt(Request.Querystring("SelGF"))
               intSelGT     = CInt(Request.Querystring("SelGT"))       

               strSelAll    = Request.Querystring("SelAll")
               strSelRec    = Request.Querystring("SelRec")
               strSelheld   = Request.Querystring("SelHeld")
               strSelExc    = Request.Querystring("SelExc")
       '        response.write("SelDojoCd=>"&strSelDojoCd&"<, len=>"&len(strSelDojoCd)&"<<br>")
        '       response.write("SelDiv1=>"&strSelDiv1&"<<br>")
       '        response.write("SelDiv2=>"&strSelDiv2&"<<br>")
      '         response.write("SelDiv3=>"&strSelDiv3&"<<br>")
      '         response.write("SelDiv4=>"&strSelDiv4&"<<br>")
             '  response.write("SelGF  =>"&intSelGF&"<<br>")
             '  response.write("SelGT  =>"&intSelGT&"<<br>")
               
     '          response.write("SelAll=>"&strSelAll&"<<br>")
      '         response.write("SelRec=>"&strSelRec&"<<br>")
       '        response.write("SelHeld=>"&strSelHeld&"<<br>")
         '      response.write("SelExc=>"&strSelExc&"<<br>")
          
               bitChooseAllDojos = 0
               strWhereSelDojoCd = ""
               intDefaultSelDojoNum = 1
               if strSelDojoCd = "all" then 
                  intDefaultSelDojoNum = 1 
                  bitChooseAllDojos = 1
               else
                  if Len(Trim(strSelDojoCd)) > 0 then 
                   '  intDefaultSelDojoNum = GetDojoNum(strSelDojoCd)
                     strWhereSelDojoCd    = " AND dms_TestPull.DojoCd = '"&strSelDojoCd&"' "
                  end if   
               end if      
              'response.write("SelDojoCd="&strSelDojoCd&", Where=>"&strWhereSelDojoCd&"<, Default=>"&intDefaultSelDojoNum&"<<br>")
              
               strWHEREDiv = ""               
               if len(strSelDiv1) = 0 then 
                  strWhereSelDiv1 = ""
               else 
                  strWhereDiv   = " (dms_TestPull.Division = '1') "
                  strdefaultDiv1 = " CHECKED "
               end if   
         '      response.write("SelDiv1="&strSelDiv1&", Where=>"&strWhereSelDiv1&"<<br>")
               

               if len(strSelDiv2) = 0 then 
                  strWhereSelDiv2 = ""
               else 
                  if len(Trim(strWhereDiv)) > 0 Then
                     strPrefix = " OR "
                  else
                     strPrefix = ""
                  end if
                  strWhereDIV   = strWhereDiv & strPrefix & " (dms_TestPull.Division = '2') "
                  strdefaultDiv2 = " CHECKED "
               end if   
       '        response.write("SelDiv2="&strSelDiv2&", Where=>"&strWhereSelDiv2&"<<br>")
               
               if len(strSelDiv3) = 0 then 
                  strWhereSelDiv3 = ""
               else 
                  if len(Trim(strWhereDiv)) > 0 Then
                     strPrefix = " OR "
                  else
                     strPrefix = ""
                  end if
               
                   strWhereDIV   = strWhereDiv & strPrefix &  " (dms_TestPull.Division = '3') "
                   strdefaultDiv3 = " CHECKED "
               end if   
        '      response.write("SelDiv3="&strSelDiv3&", Where=>"&strWhereSelDiv3&"<<br>")
               
               if len(strSelDiv4) = 0 then 
                  strWhereSelDiv4 = ""
               else 
                   if len(Trim(strWhereDiv)) > 0 Then
                     strPrefix = " OR "
                  else
                     strPrefix = ""
                  end if
                   strWhereDIV   = strWhereDiv & strPrefix &  " (dms_TestPull.Division = '4' ) "
                  strdefaultDiv4 = " CHECKED "
               end if   
               if len(Trim(strWhereDiv)) > 0 then
                  strWhereDiv = " AND ( "&strWhereDiv&" ) "
               end if
               
     '          response.write(" WhereDiv=>"&strWhereDiv&"<<br>")
               
               if intSelGF =  0 then 
                  intSelGF = 1
               end if
               if intSelGT = 0 then 
                  intSelGT = 15
               end if

               strWhereRank = " AND (NextRankID >= "&intSelGF&" AND NextRankID <= "&intSelGT&") "
 
               strStatusWhere = ""
               if strSelAll = "a" then
                  strStatusWhere = " "
                  strStatusDefaultAll = " CHECKED "
               else 
                  if strSelRec = "r" then
                  if Len(Trim(strStatusWhere)) = 0 Then
                     strPrefix = ""
                  else
                     strPrefix = " OR "
                  end if      
                     strStatusWhere = strStatusWhere & strPrefix & "  (StatusID = 4) "
                     strStatusDefaultRec= " CHECKED "
                  else 
                     strStatusWhere = strStatusWhere & ""   
                  end if
                  if strSelHeld = "h" then
                  if Len(Trim(strStatusWhere)) = 0 Then
                     strPrefix = ""
                  else
                     strPrefix = " OR "
                  end if      
                     strStatusWhere = strStatusWhere & strPrefix & "  (StatusID = 2) "
                     strStatusDefaultHeld = " CHECKED "
                  else 
                     strStatusWhere = strStatusWhere & ""   
                  end if
                  if strSelExc = "e" then
                  if Len(Trim(strStatusWhere)) = 0 Then
                     strPrefix = ""
                  else
                     strPrefix = " OR "
                  end if      
                     strStatusWhere = strStatusWhere & strPrefix & "  (StatusID = 3) "
                     strStatusDefaultExc = " CHECKED "
                  else 
                     strStatusWhere = strStatusWhere & ""   
                  end if
                  
               end if   
                if len(Trim(strStatusWhere)) > 0 then
                  strStatusWhere = " AND ( "&strStatusWhere&" ) "
               end if
              
         '      response.write("strStatusWhere=>"&strStatusWhere&" <<br>")
              
               
                            
               strBigSelWhere = strWhereSelDojoCd&" "&strWhereDiv&" "&strWhereRANK&" "&strStatusWhere&" "
             '  response.write("<BR>Big Where=>"&strBigSelWhere&"<<BR><br>")


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
                   strSortOrderCd   = Request.QueryString("s")
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
                   strTestingSchoolsStr   = ""
                   intSchoolsCnt = 0
                   for inti = 1 to len(strTestingSchools)
                      if mid(strTestingSchools,inti,1) = "1" then
                         intSchoolsCnt = intSchoolsCnt + 1
                         if Len(strTestingSchoolsWhere) > 0 then strTestingSchoolsWhere = strTestingSchoolsWhere & " OR " end if
                         if Len(strTestingSchoolsStr) > 0 then strTestingSchoolsStr = strTestingSchoolsstr & ", " end if
                        
                          if inti = 1 then 
                            strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'npk' " 
                            strTestingSchoolsStr = strTestingSchoolsStr & "New Paltz "
                         end if
                         if inti = 2 then 
                            strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'bk ' " 
                            strTestingSchoolsStr = strTestingSchoolsStr & "Brooklyn "
                         end if
                         if inti = 3 then
                            strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'kin' " 
                            strTestingSchoolsStr = strTestingSchoolsStr & "Kinnelon "
                         end if
                         if inti = 4 then 
                            strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'pv ' "
                            strTestingSchoolsStr = strTestingSchoolsStr & "Pleasant Valley "
                            
                         end if
                         if inti = 5 then 
                            strTestingSchoolsWhere = strTestingSchoolsWhere & " DojoCd = 'ef ' " 
                            strTestingSchoolsStr = strTestingSchoolsStr & "East Fishkill "

                         end if
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
                   Select Case strSortOrderCd
                          Case 0
                             strSortOrder = "ORDER BY Division, AnchorDayOrder, LName, FName, MName"
                          Case 1
                             strSortOrder = "ORDER BY LName, FName, MName, Division, AnchorDayOrder  "
                          Case 2
                             strSortOrder = "ORDER BY RankID, LName, FName, MName, Division, AnchorDayOrder  "
                          Case 3
                             strSortOrder = "ORDER BY LastTestingDate, LName, FName, MName, Division, AnchorDayOrder  "
                          Case 4
                             strSortOrder = "ORDER BY StatusID, LName, FName, MName, Division, AnchorDayOrder  "
                          Case 5
                             strSortOrder = "ORDER BY NextRankID, LName, FName, MName, Division, AnchorDayOrder  "
                          Case 6
                             strSortOrder = "ORDER BY Division, LName, FName, MName,  AnchorDayOrder  "
                          Case 7
                             strSortOrder = "ORDER BY AnchorDayOrder,division,Rankid, LName, FName, MName "
                          Case 8
                             strSortOrder = "ORDER BY DojoCd, Division, AnchorDayOrder, LName, FName, MName"
                          Case 9
                             strSortOrder = "ORDER BY AddType DESC, DojoCd, Division, AnchorDayOrder, LName, FName, MName"
                            
                          Case else
                             strSortOrder = "ORDER BY Division, AnchorDayOrder, LName, FName, MName"
End Select  
               if Session("GodMode") = 1 then
                  bitGodMode = 1
               else
                  bitGodMode = 0
               end if
               strOwnDojoCd = Session("DojoCd")
               strWHERE = ""
               strParticDojo = ""
               if bitGodMode = 1 then
                  strWhere = ""
               else
                  intpos = GetDojoNum( strOwnDojoCd )
                  strWHERE = " AND dms_TestPull.DojoCd = '"&strOwnDojoCd&"' "
               end if    
                  
                   
			   strSQLPull= "Select TestingID, dms_TestPull.StudentID , dms_TestPull.DojoCd, XDojoCd, dms_TestPull.LastTestingDate,  RankID, NextRankID, dms_TestPull.Division, dms_TestPull.Anchorday, AnchorDayOrder, StatusID, AddType, LName, FName, MName FROM dms_TestPull " 
			   strSQLPull= strSQLPull& "INNER JOIN dms_Student ON dms_Student.StudentID = dms_TestPull.StudentID "
			   strSQLPull= strSQLPull& "WHERE TestingID = '"&intTestingID&"' "&strBigSelWhere
               strSQLPull= strSQLPull& strSortorder& " ;"
			  ' response.write("strSQL=>"&strSQLpull&"<<br>") 
			   Call TopLine()
               %>

        
             
 <iframe id="MyFrame" name="MyFrame" style="display:none;"></iframe>
 <h1> Edit Rank List</h1>
 

 <h2>Sponsor: <%=strSponsoringSchoolname%></h2>		
	   
			   <table>
                  <tr>
                     <td colspan="12"> <b>Participating Schools:</b> <%=strTestingSchoolsStr%></td>
                  </tr>
			   
                  <tr>
                     <td colspan="12"> <b>Test Date:</b>  <%=datTestingDate%>&nbsp;&nbsp; <b>Test Location:</b><%=strTestingLocationName%>&nbsp;&nbsp;<b> Testing ID:</b><%=intTestingID%>&nbsp;&nbsp;&nbsp;  <a href="Testing.asp"><span style="fontweight: bold;"><i class="far fa-certificate"></i>&nbsp;Return to Testing</span></a></td>
 
                  </tr>
               <table>
               <table>
                     <form name="SelectCriteria" id = "SelectCriteria" action="SetREHCriteria.asp" method="post">
                     <% if len(trim(strMsg)) > 0 then %>
                      <tr>
                         <td class="drowerror" colspan="12"><%=strMsg%></td>
                      </tr>   
                     <% end if 
                     strSelGF=CStr(intSelGF)
                     strSelGT=CStr(intSelGT)
                      %>
                  <tr>
                      <td> <b>Dojo:</b> <% Call SelectDojo("SelDojoCd",strSelDojoCd, 1) %></td>
                      <td> <b>Div:</b><input type="checkbox" name="Seldiv1" value="1" <%=strdefaultDiv1%> >1 <td>
                      <td> <input type="checkbox" name="seldiv2" value="2" <%=strdefaultDiv2%> >2 <td>
                      <td> <input type="checkbox" name="seldiv3" value="3" <%=strdefaultDiv3%> >3 <td>
                      <td> <input type="checkbox" name="seldiv4" value="4" <%=strdefaultDiv4%> >4 <td>
                      <td> <b>From/To:</b><% Call SelectFromTableIndRange("RankCd",strSelGF, "SelGF", 2, 16) %></td>
                      <td>                <% Call SelectFromTableIndRange("RankCd",strSelGT, "SelGT", 2, 16) %></td>
                      <td> <b>Status:</b><input type="checkbox" name="selall"  value="a"  <%=strStatusDefaultAll%>>all<td>
                      <td> <input type="checkbox" name="selrec"  value="r" <%=strStatusDefaultRec%>>rec<td>
                      <td> <input type="checkbox" name="selheld" value="h" <%=strStatusDefaultHeld%>>held<td>
                      <td> <input type="checkbox" name="selexc"  value="e" <%=strStatusDefaultExc%>>exc<td>
                      <td ><input type="submit" value="Select"></td>
                      
                  </tr>
               </form>
               </table>

               
               
               <table>    
			      <tr>
	          <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			        <td class="dtabh">&nbsp;</td>
	          		<td class="dtabh">&nbsp;</td>
                    <td class="dtabh">&nbsp;</td>
                    <td class="dtabh"><a href="editpullData.asp?s=3">Last</a></td>
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
			        <td class="dtabh"><a href="editpullData.asp?s=8">Dojo</a></td>
       	            <td class="dtabh"><a href="editpullData.asp?s=1">Student Name</a></td>
                    <td class="dtabh"><a href="editpullData.asp?s=2">Last Rank</a></td>
                    <td class="dtabh"><a href="editpullData.asp?s=3">Test Date</a></td>
                    <td class="dtabh" ><a href="editpullData.asp?s=4">Status</a></td>
                    <td class="dtabh" >&nbsp;</td>
                    <td class="dtabh" ><a href="editpullData.asp?s=5">Next Rank</a></td>
                    <td class="dtabh" >&nbsp;</td>
                    <td class="dtabh" >&nbsp;</td>                  
                    <td class="dtabh"><a href="editpullData.asp?s=6">Div</a></td>
                    <td class="dtabh"><a href="editpullData.asp?s=7"><i class="fa fa-anchor"></a></td>
                    <td class="dtabh" align="center"><a href="editpullData.asp?s=9">Src</a></td>
                    <td class="dtabh">Xfer</td>

	  	  	      </tr>    
                  
			<%
			oRS.Open strSQLPull
			intCnt = 0
		    do while not ors.eof
		    intCnt = intCnt + 1
			   strLname            = ors("LName")
			   strfname            = ors("FName")
			   strMName            = ors("MName")
			   strStudentName      = left(Trim(strFName)&" "&Trim(strLName), 32)
			   intStudentID        = ors("StudentID")
			   strDojoCd           = ors("DojoCd")
			   strDivision         = ors("Division")
			   strAnchorDay        = ors("AnchorDay")
			   datLastTested       = ors("LastTestingDate")
			   strRankCd           = ors("RankId")
			   intRankID           = CInt(strRankCd)
			   strXDojoCd          = ors("XDojoCd")
			   strAddType          = ors("AddType")
			   strBGC = ""
			   if strAddType = "O" then
     			   strBgc="bgcolor='gainsboro'" 
				   strSrc = "Out "&strXDojoCd
				   strClassname="class='xferout'"
				   strspancolor = "<span style='color:blue; background-color: gainsboro; '>"
			   end if
			   if strAddType = "X" then
			      strBgc="bgcolor='mistyrose'" 
				   strSrc = "In&nbsp;"&strXDojoCd
				   strClassname="class='xferin'"
				   strspancolor = "<span style='color:blue; background-color: mistyrose; '>"


			   end if
			   if strAddType = "1" then
			      strBGC = ""
				  strSrc = ""
				  strclassname = "class='dtab'"
				  strspancolor = "<span >"

			   end if	  
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
			   if strAddType = "1"  then strDisplayDojoCd = strDojoCd end if
			   if strAddType = "O"  then strDisplayDojoCd = strDojoCd end if
			   if strAddType = "X"  then strDisplayDojoCd = strXDojoCd   end if
			     
			  
			   
			%> 
			 <tr <%=strclassname%>>
	     <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			      <form target="MyFrame" form="form<%=intStudentID%>" id="updatepullline<%=intStudentID%>" action="updatepullline.asp?ts=<%=intTestingID%>" method="post">
				  <input type="hidden" name="TDA" value="<%=datTestingDate%>" >
			      <td  <%=strclassname%>><%=spancolor%><%=strDisplayDojoCd%></span></td>
			      <td  <%=strclassname%> id="SN<%=intStudentID%>"><%=strStudentName%></td>
                  <td <%=strclassname%>>&nbsp;<%=strRankName%>&nbsp;</td>
                  <td <%=strclassname%>>&nbsp;<%=datLastTested%>&nbsp;</td>
                  <% strJS = "onchange="&chr(34)&"document.getElementById('updatepullline"&intStudentID&"').submit();"&chr(34)%>
                  <td <%=strclassname%>><% Call SelectFromTableIndJS("TestingStatus",strStatusID,intStudentID,strJS)%></td>  
       <!--         <td <%=strclassname%>><input name="recommend" type="button" value = "R" onclick="selectFunction('TestingStatus<%=intStudentID%>','4');document.getElementById('updatepullline<%=intStudentID%>').submit()">  -->
                  <td <%=strclassname%>><input name="recommend" type="button" value = "R" onclick="selectFunction('TestingStatus<%=intStudentID%>','4');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                    
									 <input name="excuse"    type="button" value = "E" onclick="selectFunction('TestingStatus<%=intStudentID%>','3');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                        <input name="Held"      type="button" value = "H" onclick="selectFunction('TestingStatus<%=intStudentID%>','2');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                        <input name="promote"   type="button" value = "P" onclick="selectFunction('TestingStatus<%=intStudentID%>','1');document.getElementById('updatepullline<%=intStudentID%>').submit()"> &nbsp; 
                  </td>
                                  
                  <td <%=strclassname%>> &nbsp;
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
                  <td <%=strclassname%>><input name="Plus"  type="button" value = "+" onclick="selectFunctionInc('NextRankID<%=intStudentID%>',<%=KRDTCountbyDiv(intD)%>);document.getElementById('updatepullline<%=intStudentID%>').submit()"></td>
                  <td <%=strclassname%>><input name="Minus" type="button" value = "-" onclick="selectFunctionDec('NextRankID<%=intStudentID%>',<%=intNextRankIdStart%>);document.getElementById('updatepullline<%=intStudentID%>').submit()"></td>                  
                  <td <%=strclassname%> align="center" ><%=strDivision%></td>
                  <td <%=strclassname%> align="center" ><%=strAnchorDay%></td>
                  <td <%=strclassname%>><input name="Div<%=intStudentID%>" type="hidden" value="<%=strDivision%>" >
                                        <input name="SID<%=intStudentID%>" type="hidden" value="<%=intStudentID%>" >
                                        <input name="Submit<%=intStudentID%>" type="submit" id="submit-button<%=intStudentID%>"  style="display: none" onclick="window.open('updatepulline.asp.asp?ts=<%=intTestingID%>', '_parent', 'width=1000,height=400,menubar=yes'); return false; ">
                                   <%=strSRC%>
                                   </td>
                  <td <%=strclassname%> align="center"><a href="TransferToTS.asp?TID=<%=intTestingID%>&SID=<%=intStudentID%>" onclick="window.open('TransferToTS.asp?TID=<%=intTestingID%>&SID=<%=intStudentID%>', 'transfer', 'width=500,height=800,scrollbars=yes,menubar=no'); return false; "><i class="fa fa-external-link-alt"></i></a></td>
			 </tr> 
		</form>  

            <%
			ors.Movenext
			loop
			
			%>
            <tr><%
                   if intCnt = 1 then strStudents = "Student" else strStudents="Students" end if %>
               <td class="dtabh">&nbsp;</td>
               <td colspan="3" class="dtabh">&nbsp;<%=intCnt%> &nbsp<%=strStudents%></td>
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
      <%
      Elapsed = StopTimer(1)
      Response.Write "Elapsed time was: " & Elapsed 
      %>
	</div>
</body>
</html>
