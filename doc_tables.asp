<html>
<head>
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
	<title>List Tables Used</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/balloon.css";</style>
	
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
	<script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/fontawesome.js"></script>

</head>   


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
    font-family: "Times Roman", Serif;	

}
body {
    font-family: "Times Roman", Serif;
    margin-right: 0%;
    margin-left:  0%;

}
.notes {
    font-size: small;
   	color: black;
    text-align:right;
}

.center { 
	width: 100%;
    border:none;
    padding:2;
    color: black;
    text-align: center;	
}
</style>



%>
<div id="contents">
<table width="100%">
   <tr>
       <td >Table & View Documentation</td>
	   <td >&nbsp;</td>
	   <td >&nbsp;</td>

   </tr>

   <tr>
       <td class="dtabh">&nbsp;</td>
       <td class="dtabh">&nbsp;</td>
       <td class="dtabh">&nbsp;</td>
	   
   
   </tr>
<%
intcntTables = 0
strSQL1 = "select name from sys.sysdatabases where dbid=db_id()"
orsA.Open strSQL1
strDBName = orsA("name")
orsA.close
set orsa= Nothing
Set oRSA = Server.CreateObject("ADODB.Recordset")
oRSA.ActiveConnection = oConn 
   %>
   <tr>
	   <td class="dtabh">DBName</td>
       <td ><%=strDBName%></td>

	   <td >&nbsp;</td>
   </tr>
   <tr>
       <td >&nbsp;</td>
       <td >&nbsp;</td>
       <td >&nbsp;</td>
   </tr>
   <tr>
       <td class="dtabh">Table</td>
	   <td class="dtabh">Created</td>
	   <td class="dtabh">Modified</td>
   </tr>
   
   
   <%


strSQL1 = "SELECT [name] ,create_date ,modify_date FROM sys.tables ORDER BY Modify_date DESC;"
orsA.open strSQL1
do while NOT orsA.EOF
   intcntTables = intCntTables + 1
   strName       = orsA("Name")
   datCreateDate = orsA("create_date")
   datModifyDate = orsA("modify_date")
   %>
   <tr>
       <td ><%=strName%></td>
	   <td ><%=datCreateDate%></td>
	   <td ><%=datModifydate%></td>
   </tr>
   
   
   <%
   orsA.MoveNext 
loop
   %>
   <tr>
       <td >&nbsp;</td>
       <td >&nbsp;</td>
       <td >&nbsp;</td>
   </tr>
      <tr>
       <td class="dtabh">View</td>
	   <td class="dtabh">Created</td>
	   <td class="dtabh">Modified</td>
   </tr>

<%
intViews = 0
orsA.close
set orsa= Nothing
Set oRSA = Server.CreateObject("ADODB.Recordset")
oRSA.ActiveConnection = oConn 
strSQL1 = "SELECT [name] ,create_date ,modify_date FROM sys.views ORDER BY Modify_date DESC;"
orsA.open strSQL1
do while NOT orsA.EOF
   intcntViews = intCntViews + 1
   strName       = orsA("Name")
   datCreateDate = orsA("create_date")
   datModifyDate = orsA("modify_date")
   %>
   <tr>
       <td ><%=strName%></td>
	   <td ><%=datCreateDate%></td>
	   <td ><%=datModifydate%></td>
   </tr>
   
   
   <%
   orsA.MoveNext 
loop
   %>
   <tr>
       <td class="dtabh">&nbsp;</td>
       <td class="dtabh">&nbsp;</td>
       <td class="dtabh">&nbsp;</td>
   </tr>
<%

%>
   <tr>
	   <td colspan='8' class="dtabh">** End of Report ***  Number of Tables <%=intcntTables%>; Views <%=intcntViews%>***</td>
   </tr>

</table>
</div>
</body>


</html>