<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/validate.inc" -->
<head>
	<title>Attendance Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
	<script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	
<style>
@media print{
    size: landscape;
    font-size: 14pt;
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
    font-family: "Arial";		

}
body {
    font-family: "Arial";
    margin-right: 0%;
    margin-left:  0%;

}
table {
    display: table;
    border-collapse: collapse;
    border-spacing: 0px;
    -webkit-border-horizontal-spacing: 0px;
    -webkit-border-vertical-spacing: 0px;
    border-color: lightgrey;
}

.roweven {
	background-color: White; 
	line-height: 4.3mm;
	  font-size: 11pt;
}
.rowodd {
	background-color: LightBlue; 
	line-height: 4.3mm;
	  font-size: 11pt;
}	
.rowtitle {
    font-family: "Arial";
    font-size: 14pt;
	
	background-color: White; 
}
.coltitle { 
	width: 100%;
    background-color: white;
    border: none;
    padding:0;
    font-family: "Times New Roman";
    font-weight: bold;	
	
    font-size: 14pt;
    color: black;
    text-align: center;	
}
.coldiv { 
	width: 33%;
    border:0;
    padding:0;
    font-size: 10pt;
    text-decoration: underline;
    font-weight: bold;	
    color: black;
    text-align: left;	
}
.colparam { 
	width: .5cm;
    border:none;
    font-weight: bold;	
    padding:0;
    font-size: 10pt;
    color: black;
    text-align: left;	
}
	
.colname { 
	width: 4cm;
    border:0;
    padding:0;
    font-size: 9pt;    
	color: black;	
}	
.colnamet { 
	width: 4cm;
    border-bottom: 1px solid black;
    padding:0;	
	font-weight: bold;	
    font-size: 10pt;    color: black;	
}
.colphone { 
	border:	width: 22mm;
 
	padding:0;
    font-size: 8pt;	color: black;
}
.colphonet { 
	width: 22mm;
    border-bottom: 1px solid black;
    font-size: 8pt;
	font-weight: bold;	
	
	color: black;
}
.colanchor { 
	text-align: center;
	width: 5%;
	border:0;
	padding:0;
    font-size: 8pt;color: black;
}

.colanchort { 
    text-align: center;
    width: 5%;
    border-bottom: 1px solid black;
    padding:0;
	font-weight: bold;	
	
    font-size: 8pt;   
    color: black;
}
.coldob { 
	width: 5%;
	border:0;
	padding:0;
  font-size: 8pt;color: black;
}

.colage { 
	text-align: right;
	width: 1%;
	border:0;
	padding:0;
    font-size: 8pt;
	color: black;
}

.coldobt { 
	width:  5%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 8pt;color: black;
	font-weight: bold;	
	
}

.colaget { 
	text-align: right;
	width: 1%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 8pt;color: black;
	color: black;
	font-weight: bold;	
	
}

.coldayeven {
	background-color: LightBlue; 
	width: 5mm;
	border: .5px solid dimgrey;
	padding:0;
	text-align: center;
	font-size:14pt;
	color: black;
}

.coldayevenblue {
background-color: White; 
width: 5mm;
border: .5px solid dimgrey;
padding:0;
	font-size:14pt;
color: black;
text-align: center;
}


.coldayodd {
background-color: Gainsboro; 
width: 5mm;
border: .5px solid dimgrey;
padding:0;
	font-size:14pt;
color: black;
text-align: center;
}

.colday {
background-color: White;
text-align: center; 
width: 4.9mm;
border: 0;
padding:0;
font-size: 8pt;
color: black;
	font-weight: normal;	

}
.colFistWeapons {
background-color: Sandybrown; 
width: 5mm;
border: .5px solid dimgrey;
padding:0;
color: black;
text-align: right;
vertical-align: text-top;
font-size: 5pt;
}

.colnotes { 
width: 4%;
border-bottom: .5px solid black;
padding:0;
font-size: 8pt;
color: black;
}
.colnotest { 
	width: 4%;
	border-bottom: 1px solid black;
	padding:0;
    font-size: 7pt;color: black;
	color: black;
	font-weight: bold;	
	
}
.colnotesnl { 
width: 4%;
border: 0;
padding:0;
font-size: 8pt;
color: black;
}
.colnotesline { 
width: 4%;
border-bottom: 1px solid black;
padding:0;
font-size:xx-small;
color: black;
}
</style>


</head>
<%  Function DoHeading(strDojoName,strReqMonthYear, strDivisionHead, strDOW, inWeekday, intMonthlen) %>
      
      <table width="100%" >
        <tr>
          <td class="coltitle"><%=strdojoname%></td>
        </tr>  
      </table>
   
    
   
      <table width="100%" >
        <tr>
          <td class="colparam">Attendance List:&nbsp; <span style="text-decoration: underline;"><%=strReqMonthYear%></span></td>
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
           <td class="colnotesnl">&nbsp;</td>
         </tr>

         <tr class="rowtitle">
           <td class="colnamet">Name</td>
           <td class="colphonet">Phone</td>
           <td class="colanchort"><i class="fa fa-anchor"></i>/Time</td>
           <td class="coldobt">DOB</td>
           <td class="colaget">Age&nbsp;</td>
           <%
           for inti=1 to intMonthLen
               stri = cstr(inti)
               stri = "0"&stri
               stri = right(stri,2)
             %><td class="colday"><%=stri%></td><%
           Next
           %>
           <td class="colnotest">Notes/Excusals</td>
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
               <td class="colname">&nbsp;</td>
               <td class="colphone">&nbsp;</td>
               <td class="colanchor">&nbsp;</td>
               <td class="coldob">&nbsp;</td>
               <td class="colage">&nbsp;</td>
<%
    bitcolodd = 1
    for inti=1 to intMonthLen
			  intm = CInt(Mid(strDOWNo,intweekday+inti-1,1))
			  if (aryClassDays(intDojoNum, intm-1) = 0) then
				 strChar = Chr(124)
			  else
				 strChar= "&nbsp;"
			  end if
	
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
               <td class="<%=strColClassName%>"><%=strChar%></td> 
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
<% Function DoLegend
	   %>
	   <tr> <td colspan="36"><image src="../images/attendancelegend2.jpg" hi></td></tr>
	   <%
End Function
%>
<%

    Function DoDataLine(bitRowParity, strName, strFormatPhone, strAnchorday, datDOB, intAge, intRN)
    
    if bitRowParity = 1 then
      strClassName = "roweven"
    else
      strClassName = "rowodd"
    end if
    strName=left(strName,26)
%>   
 
      <tr class="<%=strClassName%>">
        <td class="colname"><%=strName%> </td>
        <td class="colphone"><%=strFormatPhone%></td>
        <td class="colanchor"><%=left(strAnchorDay,1)%>&nbsp;<%=TRIM(right(" "&strAnchorday,5))%></td>
        <td class="coldob"><%=datDOB%></td>
        <td class="colage"><%=intAge%>&nbsp;</td>
      <%
           bitcolodd = 1
           for inti=1 to intMonthLen
		      bitDayOff = 0
			  intm = CInt(Mid(strDOWNo,intweekday+inti-1,1))
			  if (aryClassDays(intDojoNum, intm-1) = 0) then
				 strChar = Chr(124)
			  else
				 strChar= "&nbsp;"
			  end if
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
			  if intm=7 AND (inti>=15 AND inti <=21) and bitIsFistMember then
			     strColClassname = "colFistWeapons"
			  end if
			  if intm=7 AND (inti>=22 AND inti <=28) and bitIsOnWeapons then
			     strColClassname = "colFistWeapons"
				 strChar = CStr(intWeaponsLevel)
			  end if

      %>      <td class="<%=strColClassName%>"><%=strChar%></td><%
           Next
        bitRowParity = NOT bitRowParity
           %>
        <td class="colnotes">&nbsp;</td>

     </tr>
  
<%
    End Function
	
	Function IsDojoOpen(strDojoCd, strDOWCd)
	if strDojoCd = "npk" then
	   IsdojoOpen = 1
	end if
    if strDojoCd = "bk" then
    end if	
	END Function
	
%>

<body>
<%
strStartDate = "04/01/2017"
strReqMM = Request.Form("selMonth")
strReqDojoCd = Request.Form("selDojoCd")
strDojoCd = strReqDojoCd
intDojoNum = GetDojoID(strDojoCd)
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
intDOWMonthStart = WeekDay(intReqMMDDYYYY,1) ' Sunday is the first DOWDAYD
'response.write("1Start DOW for Month="&intDOWMonthStart&", intWeekday="&intWeekday&"<br>")
strReqMonthYear = strReqMonthName&" - "&intReqYYYY
strDOW   ="SMTWHFASMTWHFASMTWHFASMTWHFASMTWHFASMTWHFASMTWHFA"
strDOWNo ="1234567123456712345671234567123456712345671234567"
if bitIsLeapYear then
   intMonthLen = intMonthLY(intReqMM-1)
else
   intMonthLen = intMonth(intReqMM-1)
end if
strDayNames=Mid(strDOW,intDOWMonthStart,intMonthLen)
strDayNames=Mid(strDOW,intDOWMonthStart,intMonthLen)

Dim aryClassDays(16,7)
intMaxNumDojos = 16
intNumDojos = 5
for inti = 1 to intMaxNumDojos
    for intj = 0 to 6
        aryClassdays(inti, intj) = 0
    next		
next
strSQL = "SELECT DAYID, SchoolID FROM dms_Class WHERE IsActive = 1 ORDER BY SchoolID, DayID;"
orsC.Open StrSQL
intOldSchoolID = -1
do while NOT orsC.EOF
   intDayID    = orsC("DayID")
   if intDayID = 17 then intDayID = 5 end if
   if intDayID = 18 then intDayID = 7 end if
   intDayID = intDayID - 1
   intSchoolID = orsc("SchoolID")
   if intDayID < 7 then
      aryClassDays(intSchoolID, intDayID) = 1
   end if
   
   ORSC.MoveNext
loop   

'response.write("strReqDojoCd=>"&strReqDojoCd&"<<br>")
if NOT strReqDojoCd = "all" then
   strDojoSel = "(dojoCd='"&strReqDojoCd&"') AND "
else
   strDojoSel = ""
end if 
strSQL = "Select * FROM vwAttendanceReport  WHERE IsActive = 1 AND (studDojoCd = '"&strReqDojoCd&"') AND NOT (Anchorday = 'C' or Anchorday = 'V' or Anchorday = 'E' )AND (IsInstructor = 0) ORDER BY StudDojoCd, Division, AnchordayOrder, StartTime,  LName, FName;"
   

'response.write("strSQL=>"&strSQL&"<<br>")
'response.end


orsZ.open strSQL
intcnt = 0

intRowCountMax  = 40
intRowCount     = 0

intPageCount= 0

strDojoName = GetDojoName(strReqDojoCd)

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
	  
      Call FillBlankLines(intRowCount, intRowCountMax-1)
	  Call DoLegend
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
	  
      Call FillBlankLines(intRowCount, intRowCountMax-1)
	  Call DoLegend
      Call PageBreakBefore()
      bitrowparity = 1
      Call DoHeading(strDojoName,strReqMonthYear, strDivision, strDOW, inWeekday, intMonthlen)
      intRowCount = 1
	  
   end if   
   
   ' get the data  
   strDojoCode    = orsZ("StudDojoCd")
   strLName       = Trim(orsz("LName"))
   strFName       = Trim(orsZ("Fname"))
   strName        = left(strLName&", "&strFName,20)
   datDOB         = orsZ("DOB")
   DOBMonth       = Month(datDOB)
   DOBDay         = Day(datDOB)
   DOBYear        = Year(datDOB)
   DOBYear        = DOBYear Mod 100
   DOBStr         = right("0"&cstr(DobMonth),2)&"/"&right("0"&CStr(DOBDay),2)&"/"&right("00"&cstr(DobYear),2)
   datDOB         = dobStr
   strAnchorDay   = Trim(orsZ("AnchorDay"))
   'response.write("Anchor=>"&strAnchorDay&"<<br>")
   if strAnchorDay = "R" then strAnchorDay = "M" end if
   if strAnchorDay = "L" then strAnchorDay = "T" end if
   if strAnchorDay = "N" then strAnchorDay = "W" end if
   if strAnchorDay = "TH" or strAnchorDay = "D" then strAnchorDay = "H" end if 
   if strAnchorday = "Y" then strAnchorday = "F" end if 
   if strAnchorDay = "B" or strAnchorDay = "SA" then strAnchorDay = "A"
   
   
   timStartTime   = orsZ("StartTime")
   strStartTimeH = left(timStartTime,2)
   intStartTimeH = CInt(strStartTimeH)
   if intStartTimeH > 12 then intStartTimeH = intStartTimeH - 12 end if
   strStartTimeH = CStr(intStartTimeH)
   if len(strStartTimeH  < 2 ) then strStartTimeH=" "&strStartTimeH 
   strStartTimeH = right(strStartTimeH,2)
   strStartTimeM = right(left(timstartTime,5),2)
   
   strStartTime=strStartTimeH&":"&right("0"&strStartTimeM,2)
   strPhone       = orsZ("Phone")
   bitIsOnWeapons = orsz("IsOnWeapons")
   intWeaponsLevel = orsz("WeaponsLevel")
   bitIsFistMember= orsz("IsFistMember")
   strFormatPhone = "("&Left(strPhone,3)&") "&mid(strPhone,4,3)&"-"&right(strPhone,4)
   intAge         = CalcAge(datDOB, Date()) 
   if bitRowParity = 1 then
      strClassName = "roweven"
   else
      strClassName = "rowodd"
   end if
   ' write data line
   strAge = FormatNumber(intAge,1)
   if strDivision = "4" then
      strAge = ""
      datDOB = ""
   else
      datDOB = FormatDateTime(datDOB,2)   
   end if   
   Call DoDataLine(bitRowParity, strName, strFormatPhone, strAnchorday&" "&strStartTime, DobStr, strAge, intRowCount)      
   if bitDebug then response.write("Bottom>"&intRowCount&"#"&intRowCountMax&"<<br>") end if
   
   ' set up for the next
   strOldDiv = strDivision
   orsZ.MoveNext
loop
' We are done, finish up the last page
      if intRowCount >= intRowCountMax then 
         Call PageBreakBefore()
         Call DoHeading(strDojoName,strReqMonthYear, strDivision, strDOW, inWeekday, intMonthlen)
         intRowCount = 1
         bitrowparity = 1
	  end if
	  
      Call FillBlankLines(intRowCount, intRowCountMax-1)
	  Call DoLegend


%>
</body>
</html>