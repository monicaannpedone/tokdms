<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/validate.inc" -->
<head>
	<title>Dan Certificate List</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	
<style>
@media print{
    size: portrait;
    font-size: 11pt;
    margin-right: 0%;
    margin-left:  0%;
} 
.page-break-after {
	page-break-after: always !important;
}  
.page-break-before {
	page-break-before: always !important;
}  
table {
    display: table;
    border-collapse: separate;
    border-spacing: 0px;
    border-color: grey;
}

html { margin: 0 !important  
      font-family: Arial, Helvetica, sans-serif;
}
body {
    font-family: Arial, Helvetica, sans-serif;    margin-right: 0%;
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
    font-size: 15pt;	color: black;
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
    font-size: 20pt;
    color: black;
    text-align: center;	
	font-weight: bold;
 font-family: "Times New Roman", Times, serif;	
}
.colrpttitle { 
	width: 50%;
    background-color: white;
    border: none;
    padding:0;
    font-size: 14pt;
    color: black;
    text-align: center;
	font-style: italic;
    font-weight: bold;

 font-family: "Times New Roman", Times, serif;	
}

.colpageno { 
	width: 5em;
    border:0;
    padding:2;
    font-size: 12pt;
    color: black;
    text-align: left;	
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

.colname { 
	width: 25em;
    border:0;
    padding:0;
    font-size: 11pt;    color: black;	
}	
.colnamet { 
	width: 25em;
    border-bottom: 2px solid black;
    padding:0;
    font-size: 11pt;    color: black;
    font-weight: bold;

}
.colrankname { 
	width: 15em;
    border:0;
    padding:0;
    font-size: 11pt;    color: black;	
}	
.colranknamet { 
	width: 15em;
    border-bottom: 2px solid black;
    padding:0;
    font-size: 11pt;    color: black;
    font-weight: bold;

}
.coldojocd { 
	width: 4em;
	border:0;
    font-size: 11pt;	
	color: black;
}
.coldojocdt { 
	width: 4em;
	border-bottom:2px solid black;
	padding:0;
    font-size: 11pt;	color: black;
}

.coldate { 
	width: 10em;
	border:0;
	padding:0;
    font-size: 11pt;	color: black;
}
.coldatet { 
	width: 10em;
    border-bottom: 2px solid black;
    padding:0;
    font-size: 11pt;    color: black;
    font-weight: bold;
}
.coldiv { 
	width: 4em;
	border:0;
    font-size: 11pt;	
	color: black;
}
.coldivt { 
	width: 4em;
	border-bottom:2px solid black;
	padding:0;
    font-size: 11pt;	color: black;
}

.coldojocd { 
	width: 4em;
	border:0;
    font-size: 11pt;	
	color: black;
}
.coldojocdt { 
	width: 4em;
	border-bottom:2px solid black;
	padding:0;
    font-size: 11pt;	color: black;
}
.colkaino { 
	text-align: left;
	width: 20em;
	border:0;
	padding:0;
    font-size: 11pt;color: black;
	border: none;

}
.colkainot { 
    text-align: left;
    width:20em;
    border-bottom: 2px solid black;
    padding:0;
    font-size: 11pt;   
    color: black;
}
</style>


</head>
<%  
    Function DoHeading()

%>
      
      <table width="100%" >
        <tr>
          <td class="coltitle">New Paltz Karate Academy</td>
        </tr>  
	    <tr>
          <td class="colrpttitle">Dan Certificate List</td>
        </tr>  
      </table>
      <br>
      <table width="100%" >
         <thead>
         <tr class="rowtitleunderline">
           <td class="colnamet">Kai Number</td>
           <td class="colranknamet">Promotion</td>
           <td class="coldatet">Shiai Date</td>
           <td class="colkainot">Kai ID</td>
           <td class="coldivt">Division</td>
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

<%  Function DoBlankLine() %>
         <tr class=""> 
           <td class="colname">&nbsp;</td>
           <td class="colrankname">&nbsp;</td>
           <td class="coldate">&nbsp;</td>
		   <td class="colkaino">&nbsp;</td>
           <td class="coldiv">&nbsp;</td>

         </tr>


<%      
    End Function
%>

<%  Function DoFooter()  %>
           <tr class=""> 
           <td class="colnamet">&nbsp;</td>
           <td class="colranknamet">&nbsp;</td>
           <td class="coldatet">&nbsp;</td>
		   <td class="colkainot">&nbsp;</td>
           <td class="coldivt">&nbsp;</td>
           </tr>
<%
    response.write("</tbody></table>")
    response.write("Page: "&intPageCount)
    intPageCount = intPageCount + 1
     
    End Function
%>
<% Function FillBlankLines(intRC, intMax) 
       'response.write("Blank Lines From="&intRC&" to intMax="&intMax&"<br>")
       for intJ = intRC to intMax  ' Finish off the page with blank lines
           Call DoBlankLine() 
       next
   End Function
%>
<%
    Function DoDataLine()
    
    bitIsDan = NOT bitIsKyu
%>   
 
         <tr class=""> 
           <td class="colname"><%=strName%></td>
           <td class="colrankname"><%=strNextRankName%></td>
           <td class="coldate"><%=strTestingDate%></td>
		   <td class="colkaino"><%=strKaiNo%></td>
           <td class="coldiv"><%=strDivision%></td>

         </tr>
<%
    End Function
%>
<body>
<%
intTID = Request.QueryString("tid")
strSQL = "SELECT LName, FName, StudNextRankID, NextRankName, TestingDate, KaiNumber, dojoKaitemplate, Division  FROM vwFullPull2Inst WHERE TestingID = "&intTID&" and StudNextRankID >= 15 and StatusID = 4  ORDER BY  Division, LName, FName, studNextRankID;"


orsZ.open strSQL
intcnt = 0

intRowCountMax  = 47
intRowCount     = 0

intPageCount= 1
intCountActive = 0
intCountInactive = 0




Call DoHeading()  'First Header
bitDebug = 0
intRowCount = 0
intcnt = 0
do while NOT orsZ.eof
   strDivision          = orsZ("Division")
   strLName             = orsZ("LName")
   strFName             = orsZ("FName")
   strName              = left(trim(strLName)&", "&trim(strFName),26)
   strNextRankName      = orsZ("nextRankName")
   datTestingDate       = orsZ("TestingDate")
   strTestingDate       = FormatDateTime(datTestingdate,2)
   strDojoKaiTemplate   = orsz("DojoKaiTemplate")
   intKaiNumber         = orsz("KaiNumber")
   if isNumeric(intkaiNumber) AND intKaiNumber > 0 then
      strkaiNumber = CStr(intKaiNumber)
	  strKaiNo = strDojokaiTemplate&strKaiNumber
   else
      strkaiNo = "N/A"   
   end if
   
   intRowCount = intRowCount + 1
   intcnt = intcnt + 1
   if intRowCount > intRowCountMax  then 
      call DoFooter()
      Call PageBreakBefore()
      Call DoHeading()  'First Header
      intRowCount = 1
   end if   
     
   Call DoDataLine()      
   orsZ.MoveNext
loop
' We are done, finish up the last page
    call FillBlankLines(intRowCount+1, intRowCountMax) 
    call DoFooter()
	if incnt = 0 then 
	   strMSG = "No one was promoted to Shodan or above"
	else
       strMSG = ""
    end if	   
	Response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*** Total: "&intcnt&" &nbsp;&nbsp;&nbsp;&nbsp;"&strMSG)
    

%>
</body>
</html>