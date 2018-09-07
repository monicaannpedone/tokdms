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
	width: 60%;
    background-color: white;
    border: none;
    padding:0;
    font-size: x-large;
    color: black;
    text-align: center;	
    font-family: "Times Roman", Serif;	
} 
.pagenum { 
	width: 20%;
    background-color: white;
    font-size: small;
    border: none;
    padding:0;
    color: black;
    text-align: right;	
    font-size: normal;	
}
.todaysdate { 
	width: 20%;
    background-color: white;
    font-size: small;
    border: none;
    padding:0;
    color: black;
    text-align: left;	
    font-size: normal;	
}
   
.blanktitle { 
	width: 20%;
    background-color: white;
    border: none;
    padding:0;
    font-size: x-large;
    color: black;
    text-align: center;	
    font-family: "Times Roman", Serif;	
}
.blankloc { 
	width: 40%;
    border:0;
    padding:2;
    font-size: small;
    color: black;
    text-align: center;	
}
.colloc { 
	width: 40%;
    border:0;
    padding:2;
    font-size: small;
    font-weight: normal;
    color: black;
    text-align: left;	
}
.collocunderline { 
	width: 40%;
    border:0;
    padding:2;
    text-decoration: underline;
    font-weight: bold;
    font-size: small;
    color: black;
    text-align: left;	
}

.coldatet { 
	width: 20%;
    border:0;
    padding:2;
    font-size: small;
    color: black;
    text-align: center;	
    font-family: "Times Roman", Serif;	

}
.coldateunderline  { 
	width: 40%;
    border:0;
    padding:2;
    font-weight: bold;
    text-decoration: underline;
    color: black;
    text-align: center;	
}
.coldate  { 
	width: 40%;
    border:0;
    padding:2;
    font-size: small;
    font-weight: normal;
    color: black;
    text-align: center;	
}

.colunderline { 
	width: 10%;
    border:none;
     border-bottom: 1px solid black;
    padding:2;
    font-size:small;
    color: black;
    text-align: center;	
}
.colnounderline { 
	width: 10%;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: center;	
}
.coldivision { 
	width: 20em;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: center;	
}
.colname { 
	width: 40em;
    border:none;
    padding:2;
    font-family: "Lucida Console", Monospace;	
    font-size:small;
    color: black;
    text-align: left;	
}
.colother { 
	width: 25em;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: center;	
   font-family: "Lucida Console", Monospace;	
}	
.colrank { 
	width: 35em;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: center;	
   font-family: "Lucida Console", Monospace;	
}	


</style>


</head>
<%  

    Function DoHeading(strDojoName,strLocation, datTestingDate, intPage )  
    strTestingDate = FormatDateTime(datTestingDate,2) 
%>
      
      <table width="100%" >
        <tr>
          <td class="todaysdate"><span style='font-weight: bold;'>Date :</span><%=strTestingDate%></td>
          <td class="coltitle"><%=strdojoname%></td>
          <td class="pagenum"><span style='font-weight: bold;'>Page :</span><%=intPage%></td>
        </tr>  
        <tr>
          <td class="todaysdate">&nbsp;</td>
          <td class="coltitle">Student Arrival Times</td>
          <td class="pagenum">&nbsp;</td>
        </tr>  
      </table>

      <br>
      <table width="100%" >
        <tr>
          <td class="collocunderline">Held at</td>
          <td class="titleblank">&nbsp;</td>
          <td class="coldateunderline">Date of Testing </td>
        </tr>
        <tr>
          <td class="colloc"><%= strTestingLocation %></td>
          <td class="titleblank">&nbsp;</td>
          <td class="coldate"><%=strTestingDate%></td>
        </tr>

      </table>
      <br>
      <table>
         <thead>
         <tr class="rowtitle">
           <td class="coldivision">&nbsp;</td>
           <td class="colname">Student Name</td>
           <td class="colnounderline">Present</td>
           <td class="colother">Testing Time</td>
           <td class="colrank">Rank Testing For</td>
           <td class="colother">Belt Size</td>
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
       
         <tr class="">
           <td class="coldivision">&nbsp;</td>
           <td class="colname">&nbsp;</td>
           <td class="colnounderline">&nbsp;</td>
           <td class="colother">&nbsp;</td>
           <td class="colrank">&nbsp;</td>
           <td class="colother">&nbsp;</td>
         </tr> 

<%      
    End Function
%>
<% Function FillBlankLines(intRC, intMax) 
       'response.write("Blank Lines From="&intRC&" to intMax="&intMax&"<br>")
       for intJ = intRC to intMax  ' Finish off the page with blank lines
           Call DoBlankLine(bitRowParity, intJ) 
       next
   End Function
%>
<%       
    Function DoDataLine(strDiv, strName, strTestTime, strRank, strBelt)
%>   
 
         <tr class=""> 
           <td class="coldivision"><%=strDiv%></td>
           <td class="colname"><%=strName%></td>
           <td class="colunderline">&nbsp;</td>
           <td class="colother"><%=strTestTime%></td>
           <td class="colrank"><%=strRank%></td>
           <td class="colother"><%=strBelt%></td>
         </tr> 
<%
    End Function
%>
<body>
<%
intTID = Request.QueryString("tid")
strSQL = "SELECT * from vwFullPull2 WHERE TestingID = "&intTID&" AND StatusID = 4 ORDER BY TT, StudDivision, LName, FName;"

strDojoCd = strSponsoringSchoolCd





orsZ.open strSQL
intcnt = 0

intRowCountMax  = 40
intRowCount     = 0

intPageCount= 0
intOldDiv = 0
do while NOT orsZ.eof
   intcnt = intcnt + 1
   intTestingID          = orsZ("TestingID")
   datTestingDate        = orsZ("TestingDate")
   strSponsoringSchoolCd = orsZ("SponsoringSchoolCd")
   strSponsoringSchool   = Trim(orsz("SponsorDojoName"))
   
   strDojoName           = strSponsoringSchool
   intTestingLocationCd  = orsZ("TestingLocationCD")
   strTestingLocation    = Trim(orsZ("TestingLocation"))
   strTT                 = orsZ("TT")
   strTestingStatus      = orsZ("TestingStatus")
   strTestingTime        = Trim(orsZ("TestingTime"))
   strTestTime           = strTestingTime
   intDiv           = orsZ("StudDivision")
   strDojoCd        = orsZ("StudDojoCd")
   strRank          = TRIM(orsz("StudNextRank"))    
   strBeltSizeCd    = orsZ("BeltSizeCd")
   strBeltSize = strBeltSizeCd  
  
   strLName         = Trim(orsZ("LName"))
   strFName         = Trim(orsZ("FName"))
   strName          = strLName&", "&strFName
   strArrivalTime   = Trim(orsZ("AT"))
   intTestingTime   = Trim(orsZ("TT"))
   intTG            = orsZ("TG")
  
   intRowCount = intRowCount + 1

 '  response.write(intcnt&" - "&strDojoCd&"; "&strLName&", "&strFName&" Testing :"& strTestingTime&"&strIntRankID&"; "&strBeltSizeCd&"<br>")
   if intOldDiv = intDiv then
   strDiv = ""
   else
      if intcnt > 1 then 
          response.write("</table>") 
          Call PageBreakBefore()
      end if
      intPageCount = intPageCount + 1
      Call DoHeading(strDojoName, strLocation, datTestingDate, intPageCount)  'First Header
      intRowCount = 1
      intOldDiv = intDiv
      strDiv = "<span style='font-weight: bold;'>Division: </span>"&intDiv
   end if 
   if intRowCount > intRowCountMax  then 
       response.write("</table>")
       intPageCount = intPageCount + 1
       Call PageBreakBefore()
       Call DoHeading(strDojoName, strLocation, datTestingDate, intPageCount)  'First Header
       strDiv = "Division: "&intDiv
      intRowCount = 1
   end if 
    ' strRank = intRankID
    ' response.write(intcnt&" - "&strDojoCd&"; "&strLName&", "&strFName&"; Testing :"& strTestTime&"; Arrival :"&strArrivalTime&"; Rank:"&strRank&"; "&strBeltSize&"<br>")

    Call DoDataLine(strDiv, strName, strTestTime, strRank, strBeltSize)    
   orsZ.MoveNext
loop
' We are done, finish up the last page
response.write("<tr> <td colspan=4><br><br>*** end of report *** Rows read:"&intCnt&"</td></tr></table>")

%>
</body>
</html>