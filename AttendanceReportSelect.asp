<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>Select Attendance Reports</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>

</head>

<body>
<div id="main">
      <% Call TopLine() %>
	  <div id="contents">
	  <div id="OneCol">

<%
   strUserName = Session("Username")
   strSQL = "SELECT * FROM dms_pref Where username='"&strUserName&"';"
   orsZ.Open strSQL
   if orsZ.EOF then
      response.write("Userid >"&strUserName&"< not found<br>")
      response.end
   End if
   bitCanChangeDojo = orsZ("CanChangeDojo")
   strDojo          = orsZ("DojoCd")
   ' response.write("Username=>"&strUserName&"<, CanChangeDojo=>"&bitCanChangeDojo&"<, DojoCd=>"&strDojo&"<<br>")
   orsZ.Close
   set orsZ = Nothing
   
   datDate = Date()
   intMonth = Month(datDate)
   intYear  = Year(datDate)
   intMonth = intMonth + 1
   if intMonth = 13 then
      intMonth = 1
      intYear = intYear + 1
   end if   
   Select Case(intMonth)  
          case 1
               strSELECTED1 = "SELECTED"
          case 2
               strSELECTED2 = "SELECTED"
          case 3
               strSELECTED3 = "SELECTED"
          case 4
               strSELECTED4 = "SELECTED"
          case 5
               strSELECTED5 = "SELECTED"
          case 6
               strSELECTED6 = "SELECTED"
          case 7
               strSELECTED7 = "SELECTED"
          case 8
               strSELECTED8 = "SELECTED"
          case 9
               strSELECTED9 = "SELECTED"
          case 10
               strSELECTED10 = "SELECTED"
          case 11
               strSELECTED11 = "SELECTED"
          case 12
               strSELECTED12 = "SELECTED"
               
   End Select
%>
		
        <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
        <form action="AttendanceReport.asp" method="post" name="AttendanceReport.asp"  >
            <h1>Print Attendance Sheets</h1>
            
            <br><br>
            Select the Month and Year
            <% if bitCanChangeDojo then %>
			 <tr>
			   <td >Dojo</td>
			   <td class="dLabel1">
			                     <Select name="seldojocd">
			                         <option value="all"   >All Dojos</option>
			                         <option value="npk"  SELECTED >New Paltz</option>
			                         <option value="bk"  <%=strSELECTEDb%>>Brooklyn</option>
			                         <option value="kin"  <%=strSELECTEDc%>>Kinnelon</option>
			                         <option value="pv"  <%=strSELECTEDd%>>Pleasant Valley</option>
			                         <option value="ef"  <%=strSELECTEDe%>>East Fishkill</option>
			                     </Select>
			   </td>
             </tr> 
             <% else %>
                <input type="hidden" value="<%=strDojo%>" name="seldojoCd" >
             <% end if %>
             <tr>
			   <td >Month</td>
			   <td class="entry1">
			                     <Select name="selMonth">
			                         <option value=1 <%=strSELECTED1%> >January</option>
			                         <option value=2 <%=strSELECTED2%> >February</option>
			                         <option value=3  <%=strSELECTED3%> >March</option>
			                         <option value=4  <%=strSELECTED4%> >April</option>
			                         <option value=5  <%=strSELECTED5%> >May</option>
			                         <option value=6  <%=strSELECTED6%> >June</option>
			                         <option value=7  <%=strSELECTED7%> >July</option>
			                         <option value=8  <%=strSELECTED8%> >August</option>
			                         <option value=9  <%=strSELECTED9%> >September</option>
			                         <option value=10 <%=strSELECTED10%> >October</option>
			                         <option value=11 <%=strSELECTED11%> >November</option>
			                         <option value=12 <%=strSELECTED12%> >December</option>
			                     </Select>
	            </td>	
	            			   <td >Year</td>
			   <td class="entry1">
			                     
			                         <input name = "selYYYY" type="text" value="<%=intYear%>" size="4" max-length ="4">
	            </td>			   
		   
             </tr>
             
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>  
             <tr>
			   <td ><input type="submit" name="submit" value="Submit" /></td>
			   <td align="right"><input type='reset'  /></td>			   
               <td >&nbsp;</td>   
             </tr> 
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>  
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>
            <tr>
               <td>&nbsp;</td>
               <td">&nbsp;</td>
               <td>&nbsp;</td>
            </tr>
            </form>
      </TABLE>
		</div>
	</div>
   </div>
	</body>
</html>
