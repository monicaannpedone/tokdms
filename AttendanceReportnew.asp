<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Attendance Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
    <style type="text/css" media="screen">@import "../css/font-awesome.min.css";</style>  
	<script src="https://use.fontawesome.com/2ad953023c.js"></script> 
	
<style>
/*
 * Style sheet for the CSS3 specification,
 * to be used in addition to http://www.w3.org/StyleSheets/TR/W3C-{WD,PR,REC}
 */

@media print{
    size: landscape;
    font-size: small;
}   

html { margin: 0 !important  
    font-family: "Lucida Console", Monospace;		

}
body (
    
    font-family: "Lucida Console", Monospace;
    margin-right: 0%;
    margin-left:  0%
}     
div.heading1 {
	page-break-before: always;
	page-break-after: avoid;
}
div.lineitem1 {
	page-break-before: avoid;
	page-break-after: avoid;

}
.roweven {
	background-color: LightSkyBlue; 
	line-height: 75%;
}
.rowodd {
	background-color: White; 
	line-height: 75%;
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



/* Pagination */
heading {
	 page-break-after: always 
}

</style>
</head>
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

'response.write("strReqMMDDYYYY=>"&strReqMMDDYYYY&"<<br>")
if not IsDate(strReqMMDDYYYY) then
    response.write("The date >"&strReqMMDDYYYY&"< is invalid <br>")
end if    
datReqMMDDYYYY= DateValue(strReqMMDDYYYY)
intReqYYYY = Year(datReqMMDDYYYY)
bitIsLeapYear = IsLeapYearDate(intReqMMDDYYYY)
if bitIsLeapYear = 1 then
   'response.write(intReqYYYY&" is a Leap Year <br>")
else
  ' response.write(intReqYYYY&" is not a leap Year <br>")   
end if
intReqMM  =  Month(datReqMMDDYYYY)
intReqDD  =  Day(datReqMMDDYYYY)

'response.write("datReqMMDDYYYY=>"&datReqMMDDYYYY&"<<br>")
'response.write("intReqYYYY=>"&intReqYYYY&"<<br>")
'response.write("intReqMM=>"&intReqMM&"<<br>")
'response.write("intReqDD=>"&intReqDD&"<<br>")
strReqMonthName = MonthName(intReqMM)
intDOWMonthStart = WeekDay(intReqMMDDYYYY,1) ' Sunday is the first DOW
'response.write("intDOWMonthStart=>"&intDOWMonthStart&"<<br>")

strReqMonthYear = strReqMonthName&" - "&intReqYYYY
strDOW="SMTWHFASMTWHFASMTWHFASMTWHFASMTWHFASMTWHFASMTWHFA"
if bitIsLeapYear then
   intMonthLen = intMonthLY(intReqMM-1)
else
   intMonthLen = intMonth(intReqMM-1)
end if
'response.write("month=>"&intReqMM&"<, intReqYYYY=>"&intReqYYYY&"<, month len=>"&intMonthLen&"<<br>")
strDayNames=Mid(strDOW,intDOWMonthStart,intMonthLen)

'response.write("Month Length=>"&intMonthLen&"<<br>")
'response.write("Days of the Month=>"&strDayNames&"<<br>")
strDayNames=Mid(strDOW,intDOWMonthStart,intMonthLen)
strSQL = "SELECT dojoCd, IsActive, LName, FName, DOB, Division, AnchorDay, Phone FROM  dbo.dms_Student"
StrSQL = strSQL&" Left Join dms_DowCD on dms_student.AnchorDay = dms_dowcd.dowcd "
strSQL = strSQL &" WHERE  (IsActive = 1) AND (dojoCd = '"&strReqDojoCd&"') AND (IsInstructor = 0) ORDER BY dojoCd, Division, dms_dowCD.doworder, LName, FName"



orsZ.open strSQL
intcnt = 0

intLineCountMax = 40

intLineCount = 1
intPageCount = 1
strOldDiv    = ""

if trim(strDojoCd) = "bk"  then strDojoName = "Traditional Okinawan Karate of Brooklyn" end if
if trim(strDojoCd) = "kin" then strDojoName = "Traditional Okinawan Karate of Kinnelon" end if
if trim(strDojoCd) = "pv"  then strDojoName = "Traditional Okinawan Karate of Pleasant Valley" end if
if trim(strDojoCd) = "ef"  then strDojoName = "Traditional Okinawan Karate of East Fishkill" end if
if trim(strDojoCd) = "npk" then strDojoName = "New Paltz Karate Academy" end if


intYearYYYY = Year(strReqMMDDYY)
do while NOT orsZ.eof
   strDivisionHead = orsZ("Division")
   intcnt = intcnt + 1
   'response.write("+++++++++ intLineCount=>"&intLineCount&", intPageCount=>"&intPageCount&" +++++++++<<br>")
      
   strDivision = orsZ("Division")
   'response.write(" Old Div=>"&strOldDiv&"<, Div=>"&strDivision&"<<br>")
   if intLineCount = 1  or NOT (strOldDiv = strDivision) then
      strOldDiv = strDivision
      if intLineCount > 1 then
         for intJ = intLineCount to intLineCountMax
             if bitRowParity = 1 then
                strClassName = "roweven"
             else
                strClassName = "rowodd"
             end if

      %>
        <div class="lineitem">
           <tr class="<%=strClassName%>">
               <td class="colname">&nbsp;</td>
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
      %>      <td class="<%=strColClassName%>">&nbsp;</td><%
           Next
           bitrowparity = NOT bitrowparity
           %>
               <td class="colnotes">&nbsp;</td>
          </tr>
       </div> 
     <%
         Next
      else
      end if 
       %>  <div class="page-break"></div><%

      'response.write("<br><br><br><br><br><br>***************new page*************<br><br><br><br>")
   %>  
   <div class="heading"> 
      <table width="100%" >
        <tr>
          <td class="coltitle"><%=strdojoname%></td>
        </tr>  
      </table>
   </div>   
   <div class="lineitem">   
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
   </div>
      <%
      intLineCount = 5
   end if
  ' response.write("^^^^^^^^ intLineCount=>"&intLineCount&", intPageCount=>"&intPageCount&" ^^^^^^^^^<<br>")

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
   %>
   <div class="lineitem">
      <tr class="<%=strClassName%>">
        <td class="colname"><%=strName%></td>
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
   </div>  
    <% 
   'response.write("======== intLineCount=>"&intLineCount&", intPageCount=>"&intPageCount&" =========<<br>")
   intLineCount = intLineCount +1
   if intLineCount > intLineCountMax then
      intPageCount = intPageCount + 1
      intLineCount = 1
      bitrowparity = 1
   end if   
   
  
   'response.end
   strOldDiv = strDivision
   orsZ.MoveNext
loop
      if intLineCount > 1 then
         for intJ = intLineCount to intLineCountMax
             if bitRowParity = 1 then
                strClassName = "roweven"
             else
                strClassName = "rowodd"
             end if
  

      %>
        <div class="lineitem">
           <tr class="<%=strClassName%>">
               <td class="colname">&nbsp;</td>
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
      %>      <td class="<%=strColClassName%>">&nbsp;</td><%
           Next
           bitRowParity = NOT bitRowParity
           %>
               <td class="colnotes">&nbsp;</td>
          </tr>
       </div> 
     <%
         Next
      else
      end if 
%>
</tbody>
</table>
<%
'response.write("<br>The number of Records is >"&intCnt&"<<br>")



%>
</body>
</html>