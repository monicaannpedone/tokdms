<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connectmaster.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!--#include file="../include/validateSysadmin.inc"-->
<head>
	<title>Database & Table Listing </title>
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
	width: 100%;
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
	
.colname { 
	width: 16em;
    border:0;
    padding:0;
    font-size: 12pt;
    text-align: left;	
    color: black;	
}	
.colvarlength { 
	width: 16em;
    border:0;
    padding:0;
    font-size: 12pt;
    text-align: center;	
    color: black;	
}	

.colnamet { 
	width: 16em;
    border-bottom: 3px solid black;
    padding:0;
    font-size: 12pt;    color: black;
    font-weight: bold;

}
.colrank { 
	width: 14em;
	border:0;
	padding:0;
    font-size: 12pt;	color: black;
}
.colrankt { 
	width: 14em;
    border-bottom: 3px solid black;
    font-size: 12pt;	
	color: black;
}
.colrankn { 
	width: 1em;
	border-bottom:1px solid black;
	padding:0;
    font-size: 12pt;	color: black;
}
.colranknt { 
	width: 1em;
    border-bottom: 3px solid black;
    font-size: 12pt;	
	color: black;
}
.coldiv { 
	text-align: center;
	width: 1em;
	border:0;
	padding:0;
    font-size: 12pt;color: black;
}

.coldivt { 
    text-align: center;
    width: 1em;
    border-bottom: 3px solid black;
    padding:0;
    font-size: 12pt;   
    color: black;
}

.colanchor { 
	text-align: center;
	width: 1em;
	border:0;
	padding:0;
    font-size: 12pt;color: black;
}

.colanchort { 
    text-align: center;
    width: 1em;
    border-bottom: 3px solid black;
    padding:0;
    font-size: 10pt;   
    color: black;
}
.coldob { 
	width: 4em;
	border:0;
	padding:0;
  font-size: 12pt;
  color: black;
}

.colage { 
	text-align: right;
	width: 4em;
	border:0;
	padding:0;
    font-size: 12pt;
    color: black;
}

.coldobt { 
	text-align: right;
	width: 5em;
    border-bottom: 3px solid black;
	padding:0;
    font-size: 12pt;color: black;
    text-align: left;
}

.colaget { 
	text-align: right;
	width: 2em;
    border-bottom: 3px solid black;
	padding:0;
    font-size: 12pt;
    color: black;
}
.bsize{
	background-color: White; 
	width: 2em;
	border-bottom: 0px solid black;
	padding:2;
    font-size: 12pt;
    color: black;
    text-align: center;

}
.bsizet{
	background-color: White; 
	width: 2em;
	border-bottom: 3px solid black;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}

.gsize{
	background-color: White; 
	width: 2em;
	border: 1px solid dimgrey;
	padding:2;
    font-size: 12pt;color: black;
	color: black;
    text-align: center;
	
}
.gsizet{
	background-color: White; 
	width: 2em;
	border-bottom: 3px solid black;
	padding:2;
    font-size: 12pt;color: black;
	color: black;
    text-align: center;
	
}


</style>


</head>
<body>

<%  Function PageBreakBefore() %>
     <div class="page-break-before"></div>
<%  End Function %>   
   
<%  Function PageBreakAfter()  %>
     <div class="page-break-after"></div> 
<%  End Function %>      

<% Function DoHeader(strLocale, strStage, strDBName, intPageNo) %>
<table>
  <tr>
    <td class="coltitle" colspan="2">Database Listing & Table Report&nbsp; - Page <%=intPageNo%></td>
  </tr>  
  <tr>
    <td class="coltitle" colspan="2">&nbsp;</td>
  </tr>  

  <tr>
    <td class="dlabel1">Locale</td>
    <td><%=strLocale%>
  </tr>  
  <tr>
    <td class="dlabel1">Stage</td>
    <td><%=strStage%>
  </tr>  
  <tr>
    <td class="dlabel1">DBName</td>
    <td><%=strDBName%>
  </tr>  
  <tr>
    <td class="colname">&nbsp;</td>
    <td class="colname">&nbsp;</td>
  </tr>  
  <tr>
    <td class="colnamet">Table Name</td>
    <td class="colnamet">&nbsp;</td>
    <td class="colnamet">&nbsp;</td>
    <td class="colnamet">&nbsp;</td>
  </tr>
<% End Function %>

<% Function DoHeader1(strLocale, strStage, strDBName, intPageNo) %>
<table>
  <tr>
    <td class="coltitle" colspan="2">Database Listing & Table Report&nbsp; - Page <%=intPageNo%></td>
  </tr>  
  <tr>
    <td class="coltitle" colspan="2">&nbsp;</td>
  </tr>  

  <tr>
    <td class="dlabel1">Locale</td>
    <td><%=strLocale%>
  </tr>  
  <tr>
    <td class="dlabel1">Stage</td>
    <td><%=strStage%>
  </tr>  
  <tr>
    <td class="dlabel1">DBName</td>
    <td><%=strDBName%>
  </tr>  
  <tr>
    <td class="colname">&nbsp;</td>
    <td class="colname">&nbsp;</td>
  </tr>  
  <tr>
    <td class="colnamet">Table Name</td>
    <td class="colnamet">Column Name</td>
    <td class="colnamet">Data Type </td>
    <td class="colnamet">Length</td>
  </tr>
<% End Function %>
<% Function DoDataLine(strTablename) %>
  <tr>
    <td class="colname"><%=strTableName%></td>
    <td>&nbsp;</td>
  <tr>   

<% End Function %>

<% Function DoDataLineT(strTableName, strColumnName, strDataType, strLen) %>
  <tr>
    <td class="colname"><%=strTableName%></td>
    <td class="colname"><%=strColumnName%></td>
    <td class="colname"><%=strDataType%></td>
    <td class="colvarlength"><%=strLen%></td>
  <tr>   

<% End Function %>

<%
  ' Create RecordSet A and set its ActiveConnection 
strDBName = "DB_109590_dmsdev"   
strSQL = "select DISTINCT TABLE_NAME FROM Information_schema.Columns WHERE TABLE_CATALOG = '"&strDBName&"' ORDER BY TABLE_Name;"
intRowCountMax = 37
intRowCount = 0
intPageNumber = 1
Call DoHeader(strLocale, strStage, strDBName, intPageNumber)
orsA.open strSQL

do while NOT orsA.EOF
   if intRowCount > intRowCountMax then
      Call PageBreakBefore()
      response.write("</table>")
      intRowCount = 0
      intPageNumber = intPageNumber + 1
      Call DoHeader(strLocale, strStage, strDBName, intPageNumber)
   end if 
   intRowCount = intRowCount + 1
   strTablename = orsA(0)
   Call DoDataline(strTableName)
   orsA.MoveNext
Loop
orsA.Close
set orsA = Nothing
Set oRSA = Server.CreateObject("ADODB.Recordset")
oRSA.ActiveConnection = oConn 
'
'  Now lets do the Tables
'
strDBName = "DB_109590_dmsdev" 
strSQL = "select TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, Column_Name, ORDINAL_POSITION, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH "
strSQL = strSQL & "from Information_schema.Columns "
strSQL = strSQL & "where TABLE_CATALOG = '"&strDBName&"' "
strSQL = strSQL & "ORDER BY TABLE_Name, ORDINAL_POSITION;"
strOldTN = ""
  ' Create RecordSet A and set its ActiveConnection 
  
intRowCountMax = 37
intRowCount = 0
intPageNumber = 1
orsA.open strSQL
do while NOT orsA.EOF
'for inti = 0 to 6
'   response.write("orsA("&inti&") "&orsA(inti).Name&"=>"&orsA(inti)&"<<br>")
'next
   strTableName  = orsA(2) 
   if intRowCount > intRowCountMax or NOT strTablename = strOldTN then
      response.write("</table>")
      Call PageBreakBefore()      
      intRowCount = 0
      intPageNumber = intPageNumber + 1
      Call DoHeader1(strLocale, strStage, strDBName, intPageNumber)
   end if 
   if strTableName = strOldTN then
      strTableName = ""
   else 
      strTablename = orsA(2)  
      strOldTN = strTableName
   End if
    
   strColumnName = orsA(3)
   strDataType   = orsA(5)
   strLen        = orsA(6)
   'response.write("Old=>"&strOldTN&"<, Current=>"&strTableName&"< <br>")

   intRowCount = intRowCount + 1
   Call DoDataLineT(strTableName, strColumnName, strDataType, strLen)   
   orsA.MoveNext
Loop
orsA.Close
set orsA = Nothing
Set oRSA = Server.CreateObject("ADODB.Recordset")
oRSA.ActiveConnection = oConn 
response.write("</table>")
response.write("<br>****************** End of Report ***************************************")
%>
</body>
</html>