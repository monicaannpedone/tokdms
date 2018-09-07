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
	

.colranktu { 
	width: 14em;
    border-bottom: 3px solid black;
    font-size: 12pt;	
	color: black;
}
.colrankt  { 
	width: 14em;
    border-bottom: none;
    font-size: 12pt;	
	color: black;
}
.colrank   { 
	width: 14em;
    border-bottom: none;
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
	border-bottom: none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}
.bsize{
	background-color: White; 
	width: 2em;
	border-bottom:none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}

.bsizetu{
	background-color: White; 
	width: 2em;
	border-bottom: 3px solid black;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}
.qtyt{
	background-color: White; 
	width: 2em;
	border-bottom: none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}
.qty{
	background-color: White; 
	width: 2em;
	border-bottom:none;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}

.qtytu{
	background-color: White; 
	width: 2em;
	border-bottom: 3px solid black;
	padding:2;
    font-size: 12pt;
	color: black;
    text-align: center;
	
}




</style>


</head>
<body>
<%
intTID = Request.QueryString("tid")
response.write("TID=>"&intTID&"<<br>")  
strSQL = "Select NextRankID, BeltSizeCd, Count(BeltSizeCd) as BeltCount FROM dms_TestPull " 
strSQL = strSQL & "INNER JOIN dms_Student ON dms_Student.StudentID = dms_TestPull.StudentID "
strSQL = strSQL & "WHERE TestingID = '"&intTID&"' AND  "
strSQL = strSQL & " (NextRankID = 2  OR NextRankID = 3  OR NextRankID = 5   OR NextRankID = 7 OR "
strSQL = strSQL & "  NextRankID = 9  OR NextRankID = 11 OR NextRankID = 14  OR NextRankID = 15 )"
strSQL = strSQL & " GROUP BY NextRankID, BeltSizeCd "
strSQL = strSQL & " ORDER BY NextRankID, BeltSizeCd;"

response.write("strSQL=>"&strSQL&"<<br>")
orsZ.open strSQL
intcnt = 0
do while NOT orsZ.eof
   intRankID      = orsZ("NextRankID")
   strBeltSizeCd  = orsZ("BeltSizeCd")
   intBeltCount   = orsZ("BeltCount")
   intRowCount = intRowCount + 1
   intcnt = intcnt + 1
   response.write(intNextRankID&"-"&strBeltSizecd&" :: "&intBeltCount&"<br>")  
  ' strRank = GetCodeValueText("RankCd", intNextRankID)
  ' strBelt = GetCodeValueText("BeltSizeCd", strBeltSizeCd)
  ' response.write(strRank&" "&strBelt&" "&intBeltCount)  
       
   orsZ.MoveNext
loop
' We are done, finish up the last page


%>
</body>
</html>