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
'
'  Load Kata Rank Division Table
'
dim KRDTCountbyDiv(4)
dim KRDTRankID(4,30)
dim KRDTRankName(4,30)
dim KRDTRevLookupRankID(4,30)

bitDebug = 1
'
strSQL =          "SELECT RankID, DivisionID , KataID, CodeValueText AS RankName from dms_KataRankDivTable AS KRDT "
strSQL = strSQL & "INNER JOIN dms_Codes ON RankID = dms_Codes.CodeValue " 
strSQL = strSQL & "WHERE dms_Codes.CodeName = 'RankCd' AND Active = 1 AND DivisionID < 5 "
strSQL = strSQL & "ORDER BY DivisionID, RankID ;"
orsR.open strSQL
strOldDivID = 1
for inti = 1 to 4
   KRDTCountbyDiv(inti) = 0
next 
inti = 0
do while NOT orsR.EOF 
   inti = inti + 1
   intDivisionID = orsR("DivisionID")
   intD = CInt(intDivisionID)
   KRDTCountbyDiv(intD) = KRDTCountbyDiv(intD) +1
   intRankID     = orsR("RankID")
   intR = CInt(intRankID)
   strRankName   = orsR("RankName")
   if bitDebug = 1 then response.write(" KRDTCountbyDiv(intD)="& KRDTCountbyDiv(intD)&"; Div="&intDivisionID&"; RankID="&intRankID&"; RankName="&strRankName&";<br>") end if
   KRDTRankID(intD,KRDTCountbyDiv(intD))   = KRDTCountbyDiv(intD)
   KRDTRankName(intD,KRDTCountbyDiv(intD)) = strRankName
   KRDTRevLookupRankID(intD,IntR) = KRDTCountbyDiv(intD)
   orsR.MoveNext
loop
orsr.Close
set orsr = Nothing
Set oRSr = Server.CreateObject("ADODB.Recordset") 
oRSr.ActiveConnection = oConn 

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
	    if bitDebug = 1 then response.write("Rev Lookup Div "&intj&","&inti&" -- "&KRDTRevLookupRankID(intj, inti)&"<br>") end if
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
                  
                 
                   
                   
			   strSQL = "Select TestingID, dms_TestPull.StudentID ,   RankID, NextRankID, dms_TestPull.Division, dms_TestPull.Anchorday, AnchorDayOrder, StatusID, AddType, LName, FName, MName FROM dms_TestPull " 
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
			   strRankCd           = ors("RankId")
			   strRankName         = GetRankName(strRankCd)&"     "
			   ' *************************************************************************************************
			   strNextRankName     = "***"
			   intD = cint(strDivision)
	 		   intR                = Cint(ors("NextRankId"))
	 		   intNextRankID       = Cint(KRDTRevLookupRankID(intD,intR))
	 		   strNextRankName     = KRDTRankName(intD, intNextRankID)&"     "
			   datLastTested       = GetLastTestingDate(intStudentID)
			   strStatusID         = ors("StatusID")
			   response.write("<br>"&strLname&","&strFName&"-"&intStudentID&"<<br>")
			   
'///////////////////////////////////////////////////////////////////////////////////////////		
'  Set Selects to start from original next rank not the new next rank
               bitMatch = 0
               intii = 0
			   Do while  (intii < KRDTCountbyDiv(intD)) and bitMatch = 0
			       response.write("//Div="&intD&"; intii="&intii&"; strRankID="&intRankID&"; RankName="&strRankname&"; KRDTRankID("&intD&","&intii&")=>"&KRDTRankID(intD,intii)&"<<br>")
			       if cInt(KRDTRankID(intD,intii)) = Cint(strRankCd) then
			          intjj = intii + 1
			          intNextRankIDStart   = KRDTRankID(intD,intjj)
			          strNextRankNameStart = KRDTRankName(intD,intjj)
			          response.write("/\ - Current Rank "&strRankName&"; intNextRankIDStart="&intNextRankIDStart&"; strNextRankNameStart=>"&strNextRankNameStart&"<<br>")
			          bitMatch = 1
			       end if 
			       intii = intii + 1
			   loop
	

'
'//////////////////////////////////////////////////////////////////////////////////		   
			   if bitDebug = 1 then response.write(intStudentID&" "&strLName&","&strFName&" intNextRankID=>"&intNextRankID&"<; <br>") end if
			%>
			 <tr>
	     <!--  <td class="dtabedit" align="center" ><a href="EditIndStudentMVA.asp?u=<%=intStudentID%>" onclick="window.open('EditIndStudentMVA.asp?u=<%=intStudentID%>', 'newwindow', 'width=1000,height=400,menubar=yes'); return false; "> <i class="fa fa-edit"></i></a></td> -->
			      <form form="form<%=intStudentID%>" id="updatepullline<%=intStudentID%>" action="updatepullline.asp?ts=<%=intTestingID%>" method="post">

			      <td class="dtab" id="SN<%=intStudentID%>"><%=strStudentName%></td>
                  <td class="dtab">&nbsp;<%=strRankName%>&nbsp;</td>
                  <td class="dtab">&nbsp;<%=datLastTested%>&nbsp;</td>
                  <td class="dtab"><% Call SelectFromTableInd("TestingStatus",strStatusID,intStudentID)%></td>  
                  <td class="dtab"><input name="recommend" type="button" value = "R" onclick="selectFunction('TestingStatus<%=intStudentID%>','4');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                   <input name="excuse"    type="button" value = "E" onclick="selectFunction('TestingStatus<%=intStudentID%>','3');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                   <input name="Held"      type="button" value = "H" onclick="selectFunction('TestingStatus<%=intStudentID%>','2');document.getElementById('updatepullline<%=intStudentID%>').submit()">
                                   <input name="promote"   type="button" value = "P" onclick="selectFunction('TestingStatus<%=intStudentID%>','1');document.getElementById('updatepullline<%=intStudentID%>').submit()"> &nbsp; 
                  </td>
                                  
                  <td class="dtab"> &nbsp;
                     <select name = "NextRankID<%=intStudentID%>" id="NextRankID<%=intStudentID%>">
                  <%
                        intD = CInt(strDivision)
                        if bitDebug = 1 then response.write("^^^^^^^^^^^KRDTCountbyDiv("&intd&")="&KRDTCountbyDiv(intD)&"<br>") end if
                        for inti = CInt(intNextRankIDStart)  to KRDTCountbyDiv(intD)
                           if bitDebug = 1 then response.write("## KRDTRankID("&intD&","&inti&")="&KRDTRankID(intD,inti)&" nextRankID="&Cint(NextRankId)&"<br>") end if
                           if  KRDTRankID(intD,inti) = CInt(intNextRankID) then
                               strSELECTED = "SELECTED"
                           else
                               strSELECTED = ""
                           end if        
                  %>   
                           <option value="<%= inti%>" <%=strSELECTED%> ><%=KRDTRankName(intD,inti)%></option>
                  <%
                       next
                  %>
                     </select>
                  </td>
                  <td class="dtab"><input name="plus" type="button" value = "+" onclick="selectFunctionInc('NextRankID<%=intStudentID%>',<%=KRDTCountbyDiv(intD)%>);document.getElementById('updatepullline<%=intStudentID%>').submit()"></td>
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
