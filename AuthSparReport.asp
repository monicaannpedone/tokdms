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
	line-height: 75%;
}
.coltitle { 
	width: 100%;
    background-color: white;
    border: none;
    padding:0;
    font-size: large;
    color: black;
    text-align: center;	
}
.colrpttitle { 
	width: 50%;
    background-color: white;
    border: none;
    padding:0;
    font-size: 12pt;
    font-weight: bold;    
    color: black;
    text-align: center;	
}
.coldattitle { 
	width: 25%;
    background-color: white;
    border: none;
    padding:0;
    font-size: 9pt;
    font-weight: normal;    
    color: black;
    text-align: left;	
}
	
.colname { 
	width: 20%;
    border:0;
    padding:0;
  font-size: 9pt;    color: black;	
}	
.colnamet { 
	width: 20%;
    border-bottom: 1px solid black;
    padding:0;
  font-size: 10pt;    color: black;	
}

.colage { 
	text-align: center;
	width: 3%;
	border:0;
	padding:0;
    font-size: 9pt;color: black;
}

.colaget { 
	text-align: center;
	width: 3%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 10pt;color: black;
	color: black;
}
.coldiv { 
	text-align: center;
	width: 2%;
	border:0;
	padding:0;
    font-size: 9pt;color: black;
}

.coldivt { 
	text-align: center;
	width: 2%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 10pt;color: black;
	color: black;
}
.colrank { 
	text-align: center;
	width: 15%;
	border:0;
	padding:0;
    font-size: 9pt;color: black;
}

.colrankt { 
	text-align: center;
	width: 15%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 10pt;color: black;
	color: black;
}
.colcomments { 
	text-align: center;
	width: 25%;
	border-bottom: 1px solid grey;
	padding:0;
    font-size: 9pt;color: black;
}

.colcommentst { 
	text-align: center;
	width:25%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 10pt;color: black;
	color: black;
}
.colinitials { 
	text-align: center;
	width: 15%;
	border: 1px solid dimgrey;
	padding:2;
    font-size: 9pt;color: black;
}

.colinitialst { 
	text-align: center;
	width: 15%;
    border-bottom: 1px solid black;
	padding:0;
    font-size: 10pt;color: black;
	color: black;
}

.coldayeven {
	background-color: White; 
	width: 1%;
	border: 1px solid dimgrey;
	padding:2;
	font-size:xx-small;
	color: black;
}

.coldayevenblue {
background-color: LightSkyBlue; 
width: 1%;
border: 1px solid dimgrey;
padding:2;
font-size:xx-small;
color: black;
}


.coldayodd {
background-color: Gainsboro; 
width: 1%;
border: 1px solid dimgrey;
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
font-size: x-small;
color: black;
}

.colnotes { 
width: 5%;
border-bottom: .5px solid black;
padding:2;
font-size:xx-small;
color: black;
}
.colnotest { 
	width: 5%;
	border-bottom: 1px solid black;
	padding:2;
    font-size: 10pt;color: black;
	color: black;
}
.colnotesnl { 
width: 5%;
border: 0;
padding:2;
font-size:xx-small;
color: black;
}
.colnotesline { 
width: 5%;
border-bottom: 1px solid black;
padding:2;
font-size:xx-small;
color: black;
}
</style>


</head>
<%  Function DoHeading(strDojoName) 

    %>
      
      <table >
        <tr>
          <td class="coltitle"><%=strdojoname%></td>
        </tr> 
      </table>
      <table>   
        <tr>
          <td class="coldattitle"> <%=FormatDateTime(Date(),1) %> </td>
          <td class="colrpttitle">Sparring Authorization</td>
          <td class="coldattitle">&nbsp;</td>
        </tr>  
        
      </table>
   
    
      <br>
      <table width="100%" >
         <thead class="rowtitle">
           <td class="colnamet">Student Name</td>
           <td class="colaget">Age</td>
           <td class="coldivt">Division</td>
           <td class="colrankt">Rank</td>
           <td class="colinitialst">Instr Initials</td>
           <td class="colcommentst">Comments</td>
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
%>
       
      <tr class="<%=strClassName%>">
        <td class="colcomments">&nbsp;</td>
        <td class="colcomments">&nbsp;</td>
        <td class="colcomments">&nbsp;</td>
        <td class="colcomments">&nbsp;</td>
        <td class="colinitials">&nbsp;</td>
        <td class="colcomments">&nbsp;</td>

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
    Function DoDataLine(strName, intAge, strDivision, strRankName,  intRN)
    
    strName=left(strName,26)
%>   
 
      <tr class="<%=strClassName%>">
        <td class="colname"><%=strName%></td>
        <td class="colage"><%=intAge%></td>
        <td class="coldiv"><%=strDivision%></td>
        <td class="colrank"><%=strRankName%></td>
        <td class="colinitials"></td>
        <td class="colcomments"></td>

     </tr>
  
<%
    End Function
%>
<body>
<%
	' Set Up Defaults
	strusername = Session("Username")
    strSQLA = "SELECT DojoCd, CanChangeDojo, SelActive, SelInactive  FROM dms_pref WHERE username='"&strUserName&"';"
    oRSA.Open strSQLA
    if orsA.EOF then
       strErrMsg   = "Preference Record not found for user >"&strUsername&"<"
       strErrCode  = "FINDEX001"
       strMod      = "Index"
       strSev      = "F"
       RecordException strErrCode, strErrMsg, strMod, strSev 
       
    end if
    
    strDojoCd                = trim(orsA("DojoCd"))

'strSQL = "Select * FROM dms_Student  Left Join dms_Codes on dms_student.AnchorDay = dms_Codes.CodeValue  WHERE CodeName = 'AnchorDay' AND kaicd = '"&strkaicd&"' "&strDojoSel&"  "&strChoose&" ORDER BY "&strOrder&" ;"
   
strSQL = "SELECT dojoCd, IsActive, LName, FName, DOB, Division, AnchorDay, AuthSpar, RankCd FROM  dbo.dms_Student"
strSQL = strSQL &" WHERE AuthSpar = 0 AND  (IsActive = 1) AND DojoCd = '"&strDojoCd&"' AND (IsInstructor = 0) ORDER BY  Division, LName, FName, MName "
'response.write("strSQL=>"&strSQL&"<<br>")
'response.end


orsZ.open strSQL
intcnt = 0
if orsZ.EOF then
   response.write("EOF on student")
   response.end
end if   
intRowCountMax  = 50
intRowCount     = 0

intPageCount= 0
strDojoCd   = orsZ("DojoCd")
strDojoName = GetDojoName(strDojoCd)

Call DoHeading(strDojoName)  'First Header
bitDebug = 0
do while NOT orsZ.eof
   intRowCount = intRowCount + 1
   strDivision = orsZ("Division")
 
   if bitDebug then response.write("Row>"&intRowCount&"< DIV Old>"&strOldDiv&"<, New>"&strDivision&"< <<br>") end if
   intcnt = intcnt + 1
   if intRowCount > intRowCountMax  then bitPageBreak = 1 else bitpageBreak = 0 end if
   if bitDebug  then response.write("Div Break=>"&bitDivBreak&"<, PageBreak=>"&bitPageBreak&"<<br>") end if  
   if bitPageBreak then
      if bitDebug then response.write("Page Break NOT DIV Break<br>") end if
      intPageCount = intPageCount + 1
      response.write("</tbody></table>")
      Call PageBreakBefore()
      Call DoHeading(strDojoName,strReqMonthYear, strDivision, strDOW, inWeekday, intMonthlen)
      intRowCount = 1
      bitrowparity = 1
   end if   
   
   ' get the data  
   strDojoCode    = orsZ("DojoCd")
   strLName       = Trim(orsz("LName"))
   strFName       = Trim(orsZ("Fname"))
   strName        = Trim(strLName&", "&strFName)
   datDOB         = orsZ("DOB")
   bitAuthSpar    = orsZ("AuthSpar")
   strRankCd      = orsZ("RankCd")
   strRankName    = Trim(GetRankName(strRankCd))
   intAge         = CalcAge(datDOB, Date()) 
   ' write data line
   strAge = FormatNumber(intAge,1)
   if strDivision = "4" then
      strAge = ""
      datDOB = ""
   end if   
   Call DoDataLine(strName, strAge, strDivision, strRankName, intRowCount)      
   if bitDebug then response.write("Bottom>"&intRowCount&"#"&intRowCountMax&"<<br>") end if
   
   ' set up for the next
   orsZ.MoveNext
loop
' We are done, finish up the last page
Call FillBlankLines(intRowCount+1, intRowCountMax) 


%>
</body>
</html>