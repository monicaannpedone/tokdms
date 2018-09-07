<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/validate.inc" -->

	<title>Attendance Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
    <style type="text/css" media="screen">@import "../css/font-awesome.min.css";</style>  
	<script src="https://use.fontawesome.com/2ad953023c.js"></script> 
	
<style>
@media print{
    size: landscape;
    font-size: small;
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
	line-height: .25cm;
}
.rowodd {
	background-color: White; 
	line-height: .25cm;
}	
.rowtitle {
	background-color: White; 
	line-height: 75%;
}
.coltitle { 
	width: 100%;
    background-color: white;
    border: none;
    padding:2;
    font-size: large;
    color: black;
    text-align: center;	
}
.coldiv { 
	width: 33%;
    border:0;
    padding:2;
    font-size: small;
    color: black;
    text-align: center;	
}
.colparam { 
	width: 33%;
    border:none;
    padding:2;
    font-size:small;
    color: black;
    text-align: left;	
}
	
.colname { 
	width: 15%;
    border:0;
    padding:2;
    font-size:xx-small;
    color: black;	
}	
.colphone { 
	width: 9%;
	border:0;
	padding:2;
	font-size:xx-small;
	color: black;
}

.colanchor { 
text-align: center;
width: 2%;
border:0;
padding:2;
font-size:xx-small;
color: black;
}

.coldob { 
width: 6%;
border:0;
padding:2;
font-size:xx-small;
color: black;
}

.colage { 
text-align: right;
width: 3%;
border:0;
padding:2;
font-size:xx-small;
color: black;
}
.coldayeven {
background-color: White; 
width: 1%;
border: 2p solid black;
padding:2;
font-size:xx-small;
color: black;
}

.coldayevenblue {
background-color: LightSkyBlue; 
width: 1%;
border: 2p solid black;
padding:2;
font-size:xx-small;
color: black;
}


.coldayodd {
background-color: Gainsboro; 
width: 1%;
border: 2p solid black;
padding:2;
font-size:xx-small;
color: black;
}

.colday {
background-color: White;
text-align: center; 
width: 1%;
border: 0;
padding:2;
font-size:xx-small;
color: black;
}

.colnotes { 
width: 5%;
border: 0;
padding:2;
font-size:xx-small;
color: black;

.colnotesline { 
width: 5%;
border-bottom: 2px solid black;
padding:2;
font-size:xx-small;
color: black;
</style>

<head>
</head>
<%  Function DoHeading(strDojoName,strReqMonthYear, strDivisionHead, strDOW, inWeekday, intMonthlen) %>
      
      <table width="100%" >
        <tr>
          <td class="coltitle"><%=strdojoname%></td>
        </tr>  
      </table>
   
    
      <br>
      <table width="100%" >
        <tr>
          <td class="colparam">Attendance List:<%=strReqMonthYear%></td>
          <td class="coldiv">Division <%=strDivisionHead%></td>
          <td class="colparam">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" >
         <thead>
         <tr class="rowtitle">
           <td class="colname">&nbsp;</td>
           <td class="colphone">&nbsp;</td>
           <td class="colanchor">&nbsp;</td>
           <td class="coldob">&nbsp;</td>
           <td class="colage">&nbsp;</td>
           <%
           for inti=1 to intMonthLen
             %><td class="colday"><%=Mid(strDOW,intWeekday+inti-1,1)%></td><%
           Next
           %>
           <td class="colnotes">&nbsp;</td>
         </tr>

         <tr class="rowtitle">
           <td class="colname">Name</td>
           <td class="colphone">Phone</td>
           <td class="colanchor"><i class="fa fa-anchor"></td>
           <td class="coldob">DOB</td>
           <td class="colage">Age</td>
           <%
           for inti=1 to intMonthLen
               stri = cstr(inti)
               stri = "0"&stri
               stri = right(stri,2)
             %><td class="colday"><%=stri%></td><%
           Next
           %>
           <td class="colnotes">Notes/Excusals</td>
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

<%  Function DoBlankLine(bitRowParity, intRN)
    if bitRowParity = 1 then
       strClassName = "roweven"
    else
       strClassName = "rowodd"
    end if
%>
       
           <tr class="<%=strClassName%>">
               <td class="colname"><%=intRN%>&nbsp;</td>
               <td class="colphone">&nbsp;</td>
               <td class="colanchor">&nbsp;</td>
               <td class="coldob">&nbsp;</td>
               <td class="colage">&nbsp;</td>
<%
    bitcolodd = 1
    for inti=1 to intMonthLen
        if bitcolodd then
           bitcolodd = 0
           strColClassName = "coldayodd"
        else
           if bitrowparity = 1 then
              strColClassName = "coldayevenblue"
           else
              strColClassName = "coldayeven"
           end if         
           bitcolodd = 1
        end if   
%>   
               <td class="<%=strColClassName%>">&nbsp;</td> 
<%   Next
          bitRowParity = NOT bitRowParity
%>
               <td class="colnotes">&nbsp;</td>
           </tr>
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
    Function DoDataLine(bitRowParity, strName, strFormatPhone, strAnchorday, datDOB, intAge, intRN)
    
    if bitRowParity = 1 then
      strClassName = "roweven"
    else
      strClassName = "rowodd"
    end if
      
%>   
 
      <tr class="<%=strClassName%>">
        <td class="colname"><%=intRN%>-<%=strName%></td>
        <td class="colphone"><%=strFormatPhone%></td>
        <td class="colanchor"><%=strAnchorDay%></td>
        <td class="coldob"><%=datDOB%></td>
        <td class="colage"><%=intAge%></td>
      <%
           bitcolodd = 1
           for inti=1 to intMonthLen
              if bitcolodd then
                 bitcolodd = 0
                 strColClassName = "coldayodd"
              else
                 if bitrowparity = 1 then
                     strColClassName = "coldayevenblue"
                 else
                     strColClassName = "coldayeven"
                 end if         
                 bitcolodd = 1
              end if   
      %>      <td class="<%=strColClassName%>">&nbsp;</td><%
           Next
        bitRowParity = NOT bitRowParity
           %>
        <td class="colnotesline">&nbsp;</td>

     </tr>
  
<%
    End Function
%>
<body>
<%



strStartDate = "04/01/2017"
strReqMM = Request.Form("selMonth")
strReqDojoCd = Request.Form("selDojoCd")
strDojoCd = strReqDojoCd
strReqYYYY   = Request.Form("selYYYY")
'response.write("strReqYYYY=>"&strReqYYYY&"<<br>")
bitrowparity = 1
bitcolodd = 1
intMonth   = array(31,28,31,30,31,30,31,31,30,31,30,31)
intMonthLY = array(31,29,31,30,31,30,31,31,30,31,30,31) 
strReqMMDDYYYY = strReqMM&"/01/"&strReqYYYY
intWeekday=Weekday(strReqMMDDYYYY,1)

if not IsDate(strReqMMDDYYYY) then
    response.write("The date >"&strReqMMDDYYYY&"< is invalid <br>")
end if    
datReqMMDDYYYY= DateValue(strReqMMDDYYYY)
intReqYYYY = Year(datReqMMDDYYYY)
bitIsLeapYear = IsLeapYearDate(intReqMMDDYYYY)
intReqMM  =  Month(datReqMMDDYYYY)
intReqDD  =  Day(datReqMMDDYYYY)

strReqMonthName = MonthName(intReqMM)
intDOWMonthStart = WeekDay(intReqMMDDYYYY,1) ' Sunday is the first DOW

strReqMonthYear = strReqMonthName&" - "&intReqYYYY
strDOW="SMTWHFASMTWHFASMTWHFASMTWHFASMTWHFASMTWHFASMTWHFA"
if bitIsLeapYear then
   intMonthLen = intMonthLY(intReqMM-1)
else
   intMonthLen = intMonth(intReqMM-1)
end if
strDayNames=Mid(strDOW,intDOWMonthStart,intMonthLen)
strDayNames=Mid(strDOW,intDOWMonthStart,intMonthLen)
strSQL = "SELECT dojoCd, IsActive, LName, FName, DOB, Division, AnchorDay, Phone FROM  dbo.dms_Student"
StrSQL = strSQL&" Left Join dms_DowCD on dms_student.AnchorDay = dms_dowcd.dowcd "
strSQL = strSQL &" WHERE  (IsActive = 1) AND (dojoCd = '"&strReqDojoCd&"') AND (IsInstructor = 0) ORDER BY dojoCd, Division, dms_dowCD.doworder, LName, FName"



orsZ.open strSQL
intcnt = 0

intRowCountMax  = 40
intRowCount     = 0

intPageCount= 0

if trim(strDojoCd) = "bk"  then strDojoName = "Traditional Okinawan Karate of Brooklyn" end if
if trim(strDojoCd) = "kin" then strDojoName = "Traditional Okinawan Karate of Kinnelon" end if
if trim(strDojoCd) = "pv"  then strDojoName = "Traditional Okinawan Karate of Pleasant Valley" end if
if trim(strDojoCd) = "ef"  then strDojoName = "Traditional Okinawan Karate of East Fishkill" end if
if trim(strDojoCd) = "npk" then strDojoName = "New Paltz Karate Academy" end if


intYearYYYY = Year(strReqMMDDYY)
Call DoHeading(strDojoName,strReqMonthYear, 1, strDOW, inWeekday, intMonthlen)  'First Header
strOldDiv    = "1"
bitDebug = 0
do while NOT orsZ.eof
   intRowCount = intRowCount + 1
   strDivision = orsZ("Division")
 
   if bitDebug then response.write("Row>"&intRowCount&"< DIV Old>"&strOldDiv&"<, New>"&strDivision&"< <<br>") end if
   intcnt = intcnt + 1
   if NOT (strOldDiv = strDivision) then bitDivBreak  = 1 else bitDivBreak  = 0 end if
   if intRowCount > intRowCountMax  then bitPageBreak = 1 else bitpageBreak = 0 end if
   if bitDebug  then response.write("Div Break=>"&bitDivBreak&"<, PageBreak=>"&bitPageBreak&"<<br>") end if  
   if bitPageBreak and NOT bitDivBreak then
      if bitDebug then response.write("Page Break NOT DIV Break<br>") end if
      intPageCount = intPageCount + 1
      response.write("</tbody></table>")
      Call PageBreakBefore()
      Call DoHeading(strDojoName,strReqMonthYear, strDivision, strDOW, inWeekday, intMonthlen)
      intRowCount = 1
      bitrowparity = 1
   end if   
   if bitDivBreak and NOT bitPageBreak then  ' it is just a Division Break
      if bitDebug then response.write("Div Break NOT Page Break<br>") end if
      Call FillBlankLines(intRowCount, intRowCountMax) 
      Call PageBreakBefore()
      bitrowparity = 1
      Call DoHeading(strDojoName,strReqMonthYear, strDivision, strDOW, inWeekday, intMonthlen)
      intRowCount = 1
   end if 
   if bitPageBreak AND bitDivBreak then 
      if bitDebug then response.write("Page Break AND  div Break<br>") end if
      intRowCount = 1
      bitrowparity = 1
      response.write("</tbody></table>")
      Call PageBreakBefore()
      Call DoHeading(strDojoName,strReqMonthYear, strDivision, strDOW, inWeekday, intMonthlen)
   end if   
   
   ' get the data  
   strDojoCode    = orsZ("DojoCd")
   strLName       = Trim(orsz("LName"))
   strFName       = Trim(orsZ("Fname"))
   strName        = strLName&", "&strFName
   datDOB         = orsZ("DOB")
   strAnchorDay   = orsZ("AnchorDay")
   strPhone       = orsZ("Phone")
   strFormatPhone = "("&Left(strPhone,3)&") "&mid(strPhone,4,3)&"-"&right(strPhone,4)
   intAge         = CalcAge(datDOB, Date()) 
   if bitRowParity = 1 then
      strClassName = "roweven"
   else
      strClassName = "rowodd"
   end if
   if intAge > 100 then
      intAge=""
      datDOB=""
   end if 
   ' write data line
   Call DoDataLine(bitRowParity, strName, strFormatPhone, strAnchorday, datDOB, intAge, intRowCount)      
   if bitDebug then response.write("Bottom>"&intRowCount&"#"&intRowCountMax&"<<br>") end if
   
   ' set up for the next
   strOldDiv = strDivision
   orsZ.MoveNext
loop
' We are done, finish up the last page
Call FillBlankLines(intRowCount+1, intRowCountMax) 


%>
</body>
</html>