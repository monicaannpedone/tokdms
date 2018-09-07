<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/validate.inc" -->
<head>
	<title>Belt Order</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	
<style>
@media print{
    size: portrait;
    font-size: 12pt;
    margin-right: 0%;
    margin-left:  0%;
} 
.page-break-after {
	page-break-after: always !important;
}  
.page-break-before {
	page-break-before: always !important;
}  

html { margin: 0 !important  
    font-family: "Lucida Console", Monospace;		

}
body {
    font-family: "Lucida Console", Monospace;
    margin-right: 0%;
    margin-left:  0%;

}
.roweven {
	background-color: LightSkyBlue; 
	line-height: .3cm;
	  font-size: 12pt;
}
.rowodd {
	background-color: White; 
	line-height: .3cm;
	  font-size: 12pt;
}	
.rowtitle {
	background-color: White; 
    font-size: 12pt;	color: black;
    font-weight: bold;
    

}
.rowtitleunderline {
	background-color: White; 
	border-bottom: 2px solid black;
    font-size: 12pt;	color: black;
    font-weight: bold;
}
.coltitle { 
	width: 50%;
    background-color: white;
    border: none;
    padding:0;
    font-size: large;
    color: black;
    text-align: center;	
}
.coldiv { 
	width: 33%;
    border:0;
    padding:2;
    font-size: large;
    color: black;
    text-align: center;	
}
.colparam { 
	width: 33%;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: right;	
}
.colparam1 { 
	width: 33%;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: left;	
}
	

.ranktu { 
	width: 14em;
    border-bottom: 3px solid black;
    font-size: 12pt;	
	color: black;
	font-weight: bold;
	
}
.rankt  { 
	width: 14em;
    border-bottom: none;
    font-size: 12pt;	
	color: black;
	border-bottom: 3px solid black;
	font-weight: bold;
	
}
.rank   { 
	width: 14em;
    border-bottom: none;
    font-size: 12pt;	
	color: black;
	font-weight: normal;
}

.coldiv { 
	text-align: center;
	width: 1em;
	border:0;
	padding:0;
    font-size: 12pt;color: black;
}
.dojotu { 
	text-align: center;
	width: 10em;
	border:0;
	padding:0;
    font-size: 12pt;color: black;
	border-bottom: 3px solid black;
	font-weight: bold;
}
.dojot  { 
	text-align: center;
	width: 10em;
	border:0;
	padding:0;
    font-size: 12pt;color: black;
	font-weight: bold;
    
}
.dojo   { 
	text-align: center;
	width: 10em;
	border:0;
	padding:0;
    font-size: 12pt;color: black;
    font-weight: normal;
}

.bsizetu{
	background-color: White; 
	width: 10em;
	border-bottom: 3px solid black;
	padding:2;
    font-size: 12pt;
    color: black;
    text-align: center;
    font-weight:bold;
}
.bsizet{
	background-color: White; 
	width: 10em;
	border-bottom: none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	font-weight: bold;
}
.bsize{
	background-color: White; 
	width: 10em;
	border-bottom:none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
    font-weight: normal;
	
}

.qtyt{
	background-color: White; 
	width: 2em;
	border-bottom: none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	font-weight: bold;
	
}
.qty{
	background-color: White; 
	width: 2em;
	border-bottom:none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	font-weight: normal;

}

.qtytu{
	background-color: White; 
	width: 2em;
	border-bottom: 3px solid black;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	font-weight: bold;
	
}




</style>


</head>
<%  
    Function DoHeading(strDojoName,datTestingDate, intTID) 
    strTestingDate = FormatDateTime(datTestingDate,1)
   

%>
      
      <table width="100%" >
        <tr>
          <td class="coltitle"><%=strdojoname%></td>
        </tr>  
      </table>
      <br>
      <table width="100%" >
        <tr>
          <td class="colparam1">Test Date:<%=strTestingDate%>&nbsp;&nbsp; Testing ID: <%=intTID%></td>
          <td class="coldiv">Belt Ordering List</td>
          <td class="colparam">Printed: <%=Now()%></td>
        </tr>
      </table>
      <br>
      <table  >
         <thead>
         <tr class="rowtitle">
           <td class="">&nbsp; </td>
           <td class="">&nbsp; </td>
           <td class="bsizet">Belt</td>
           <td class="">&nbsp;</td>
         </tr>

         <tr class="rowtitleunderline">
           <td class="dojotu">Dojo</td>
           <td class="ranktu">Belt</td>
           <td class="bsizetu">Size</td>
           <td class="qtytu">Qty</td>
          
         </tr>
         </thead>
         <tbody>
         
<%  End Function %>

<%  Function PageBreakBefore() %>
     <div class="page-break-before"></div>
<%  End Function %>   
   
<%  Function PageBreakAfter()  %>
     <div class="page-break-after"></div> 
<%  End Function %>      

<%  Function DoBlankLine(bitIsKyu, intRN)
    bitIsDan = NOT bitIsKyu
%>
       
         <tr class="rowtitle">
           <td class="colnamet">&nbsp;</td>
           <td class="colrankt">&nbsp;</td>
           <td class="coldobt">&nbsp;</td>
           <td class="colrank">&nbsp;</td>
           

<%      
    End Function
%>
<% Function FillBlankLines(intRC, intMax) 
       'response.write("Blank Lines From="&intRC&" to intMax="&intMax&"<br>")
       for intJ = intRC to intMax  ' Finish off the page with blank lines
           Call DoBlankLine(bitRowParity, intJ) 
           'response.write("RN=>"&intJ&"<<br>")
           if intJ = intMAX then
              response.write("</tbody></table>")
           end if   
       next
   End Function
%>
<%
    Function DoDataLine(strDojoCd, strRank, strBeltSize, intQty)
    
    bitIsDan = NOT bitIsKyu
%>   
 
         <tr class="">
           <td class="dojo"><%=strDojoCd%></td>
           <td class="rank"><%=strRank%></td>
           <td class="bsize"><%=strBeltSize%></td>
           <td class="qty"><%=intqty%></td>
         </tr>
<%
    End Function
%>
<body>
<%
intTID = Request.QueryString("tid")
DIM TimeAry(12)
intNumTimes = 0
strSQLT = "SELECT CodeValueText from dms_Codes WHERE CodeName='Times' ORDER BY ValueOrder; "
orsB.open strSQLT
do while NOT orsB.EOF
   intNumTimes = intNumTimes + 1
   TimeAry(intNumTimes-1) = orsB("CodeValueText")
   orsB.MoveNext
loop  
orsB.Close
set orsB = Nothing
Set oRsB = Server.CreateObject("ADODB.Recordset") 
oRSB.ActiveConnection = oConn

strSQLts = "SELECT * FROM dms_TestSession WHERE TestingID = "&intTID&";"
orsB.open strSQLts
If  NOT orsB.EOF then
 ' response.write("Testing ID "&intTID&" Found<br>")   
  intTestingID          = orsB("TestingID")
  datTestingDate        = orsB("TestingDate")
  strSponsoringSchoolCd = orsB("SponsoringSchoolCd")
  strSponsoringSchool   = GetDojoShortName(strSponsoringSchoolCd)
  intTestingLocationCd  = orsB("TestingLocationCD")
  strTestingLocation    = GetCodeValueText("TestLocation", intTestingLocationCd)
  strTestingStatus      = orsB("TestingStatus")
else
  response.write("Testing ID "&intTID&" Not Found<br>")   
end if
               
strSQL =          "SELECT B.StudentID, LName, FName, FromRankID, BeltSizeCd, NextRankID, ToRankId, B.DojoCd as Dojo, B.Division as intDiv1, TG.Division as DCd, CAST(B.Division AS INT) as intD, "
strSQL = strSQL & "SUBSTRING(TG.Division, CAST(B.Division AS INT), 1) AS Divs,  TestingGroup,  B.[Division] as strdiv, TestTime as TT, ArrivalTime as AT,  B.TestingID, StatusID    "
strSQL = strSQL & "FROM dms_TestGroup  AS TG "
strSQL = strSQL & "INNER JOIN "
strSQL = strSQL & "(SELECT J.StudentID, LName, FName, BeltSizeCd, NextRankID, J.DojoCd, T.Division, TestingID , StatusID  "
strSQL = strSQL & " FROM dms_TestPull  AS T "
strSQL = strSQL & " INNER JOIN  "
strSQL = strSQL & "  (SELECT StudentID, LName, FName, S.DojoCd, BeltSizeCd "
strSQL = strSQL & "    FROM dms_Student  AS S  "
strSQL = strSQL & "      INNER JOIN dms_Dojo  AS D "
strSQL = strSQL & "           ON S.dojoCd  = D.dojocd) AS J "
strSQL = strSQL & "      ON T.StudentID =  J.StudentID)  AS B "
strSQL = strSQL & " ON TG.TestingID = B.TestingID and StatusID = 4 "
strSQL = strSQL & "WHERE B.TestingID = "&intTID&" and SUBSTRING(TG.Division, CAST(B.Division AS INT), 1) = '1' "
strSQL = strSQL & "ORDER BY TestingGroup, B.Division, LName, FName;    "

strDojoCd = strSponsoringSchoolCd




if trim(strDojoCd) = "bk"  then strDojoName = "Traditional Okinawan Karate of Brooklyn" end if
if trim(strDojoCd) = "kin" then strDojoName = "Traditional Okinawan Karate of Kinnelon" end if
if trim(strDojoCd) = "pv"  then strDojoName = "Traditional Okinawan Karate of Pleasant Valley" end if
if trim(strDojoCd) = "ef"  then strDojoName = "Traditional Okinawan Karate of East Fishkill" end if
if trim(strDojoCd) = "npk" then strDojoName = "New Paltz Karate Academy" end if


intYearYYYY = Year(strReqMMDDYY)
'%  Function DoHeading(bitIsKyu, strDojoName,datTestingDate, strDivisionHead, strDOW, inWeekday, intMonthlen) 


orsZ.open strSQL
intcnt = 0

intRowCountMax  = 35
intRowCount     = 0

intPageCount= 0
'response.write(intTID&"  EOF = "&orsZ.eof&"<br>")
strOldDojoCd = ""
do while NOT orsZ.eof
   strDojoCd        = orsZ("DojoCd")
   intNextRankID    = orsZ("NextRankID")
   strBeltSizeCd    = orsZ("BeltSizeCd")
   strLName         = Trim("LName")
   strFName         = Trim("FName")
   strBeltSize      = GetCodeValueText("BeltSizeCd", strBeltSizeCd)
   intArrivalTimeCd = orsZ("AT")
   if isNull(intArrivalTimeCd) then
      intArrivalTimeCd = 0
   end if   
   strArrivalTime   = TimeAry(CInt(intArrivalTimeCd))
   intTestTimeCd    = orsZ("TT")

   if isNull(intTestTimeCd) then
      intTestTimeCd = 0
   end if   
   strTestTime      = TimeAry(CInt(intTestTimeCd))
   intTG            = orsZ("TestingGroup")
   intcnt = intcnt + 1

 '  response.write(intcnt&" - "&strDojoCd&"; "&strLName&", "&strFName&" Testing :"& strTestingTime&"&strIntRankID&"; "&strBeltSizeCd&"<br>")
   if strOldDojoCd = strDojoCd then
   else
      if intcnt > 1 then 
         response.write("</table>") 
         Call PageBreakBefore()
      end if
      Call DoHeading(strDojoName, datTestingDate, intTID)  'First Header
      intRowCount = 1
      strOldDojoCd = strDojoCd
   end if 
   if intRowCount > intRowCountMax  then 
      response.write("</table>")
      Call PageBreakBefore()
      Call DoHeading( strDojoName, datTestingDate, intTID)  'First Header
      intRowCount = 1
   end if 
   strRank = GetCodeValueText("RankCd", intRankID)
    ' strRank = intRankID
   if len(trim(strBeltSizeCd)) = 0 then
      strBeltSizeCd = "Unk"
   end if   
   strBelt     = GetCodeValueText("BeltSizeCd", strBeltSizeCd)
   strRankName = GetCodeValueText("RankCd",     strNextRankID)
   response.write(intcnt&" - "&strDojoCd&"; "&strLName&", "&strFName&" Testing :"& strTestingTime&"&strRankName&"; "&strBelt&"<br>")

  ' Call DoDataLine(GetDojoShortName(strDojoCd), strRank, strBelt, intBeltCount)      
   orsZ.MoveNext
loop
' We are done, finish up the last page
response.write("<tr> <td colspan=4><br><br>*** end of report *** Rows read:"&intCnt&"</td></tr></table>")

%>
</body>
</html>