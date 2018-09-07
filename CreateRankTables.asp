<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Attendance Form</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
<script>
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
      <% Call Topline() %>
	  <div id="contents">
	  <div id="OneCol">
      <br/>
<%                
 '
'  Load Kata Rank Division Table
'

dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.CreateTextFile("c:\rank.inc",true)
f.WriteLine("' ** Rank Table **")

f.WriteLine("dim KRDTCountbyDiv(4)")
f.WriteLine("dim KRDTRankID(4,30)")
f.WriteLine("dim KRDTRankName(4,30)")
f.WriteLine("dim KRDTRankIDRev(4,30)")
dim CountbyDiv(4)")
dim RankID(4,30)")
dim RankName(4,30)")
dim RankIDRev(4,30)")
response.end
'
strSQL =          "SELECT RankID, DivisionID , KataID, CodeValueText AS RankName from dms_KataRankDivTable AS KRDT "
strSQL = strSQL & "INNER JOIN dms_Codes ON RankID = dms_Codes.CodeValue " 
strSQL = strSQL & "WHERE dms_Codes.CodeName = 'RankCd' AND Active = 1 AND DivisionID < 5 "
strSQL = strSQL & "ORDER BY DivisionID, RankID ;"
orsR.open strSQL
strOldDivID = 1
for inti = 1 to 4
   CountbyDiv(inti) = 0
next 
do while NOT orsR.EOF 
   intDivisionID = orsR("DivisionID")
   intD = CInt(intDivisionID)
   intRankID     = orsR("RankID")
   intR = CInt(intRankID)
   f.writeline("KRDTRankID("&intD&","&intR&") = 
   strRankName   = orsR("RankName")
    response.write(intDivisionID&" "&intRankID&" "&strRankName&"<br>")
   CountbyDiv(intD) = CountbyDiv(intD) +1
   RankID(intD,CountbyDiv(intD))   = intR
   RankName(intD,CountbyDiv(intD)) = strRankName
   orsR.MoveNext
loop
orsr.Close
set orsr = Nothing
Set oRSr = Server.CreateObject("ADODB.Recordset") 
'
f.Close
set f=nothing
set fs=nothing
'   KRDTCountByDiv(1) = 14

%>
		</div>	
      </div>
	</div>
</body>
</html>
