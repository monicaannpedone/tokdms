<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>Edit Student</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/regular.js"></script>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	<script>
	var selectobject=document.getElementById("AnchorClass")
	var Divobj      = document.getElementByID("Division")
	
  for (var i=0; i<selectobject.length; i++){
  if (selectobject.options[i].value == 'A' )
     selectobject.remove(i);
  }
	</script>
</head>
<body >
<div id="main">
<%
			dim DayName(9)
			dayName(1) = "Sun"
			dayname(2) = "Mon"
			dayName(3) = "Tue"
			dayname(4) = "Wed"
			dayName(5) = "Thu"
			dayname(6) = "Fri"
			dayName(7) = "Sat"
			dayname(8) = "Flo"

			%>
	  <div id="contents">
	  <div id="OneCol">
         <Table> 
          <form action="maintStudent222.asp?m=<%=strMode1%>" method="post">
	  
               <tr>
                  <td class="dtabh">Division</td>
                  <td>
                   <% 
                      Call SelectFromTableIndJS("Division", strDivision, "Division","onclick=window.alert('sometext')")				   %>
                  </td>
                  <td class="drowerror"><%=aryErrMsg8%></td>
               </tr>
         <!--   <tr> -->
          <!--       <td class="dtabh">Anchor Day</td> -->
           <!--        <td> -->
          <!--            <% Call SelectFromTable("Anchorday", strAnchorDay)%> -->
           <!--          </td> -->
           <!--        <td class="drowerror"> <%=aryErrMsg9%></td> -->
		  <!--      </tr> -->
               <tr>			   
				  <td class="dtabh">Select Anchor Class</td>
				  <td>
				    <select name="AnchorClass" id="AnchorClass" >
				  <%
				     strSQL = "SELECT ClassID, DayID, DivisionID, StartTime , ClassTypeID FROM dms_Class where DojoCd = '"&Session("DojoCd")&"' and ClassTypeID = 1 ORDER BY DivisionID, DAYID, StartTime;"
				     orsH.Open strSQL
					 do while NOT orsH.EOF
					 strSelected = ""
					    intClassID    = orsH("ClassID")
						intdayID      = orsH("dayID")
						intDivisionID = orsH("DivisionID")
						timStartTime  = orsH("StartTime")
						strDayName    = DayName(intDayID)
						if intClassID = intAnchorClass then strSelected = "SELECTED" end if
					    %><option value="<%=intClassID%>" <%=strSELECTED%>><%=strDayName%>&nbsp;Div&nbsp;<%=intDivisionID%>@<%=left(timStartTime,5)%></option><%
					    orsH.Movenext
					 loop
				     orsH.Close
                     set orsH = Nothing
                     Set oRSH = Server.CreateObject("ADODB.Recordset") 
                     oRSH.ActiveConnection = oConn 
					 
				  %></select>
				  </td>
				  <td class="drowerror"> <%=aryErrMsg23%></td>
               </tr> 
           </form>
          </TABLE>
			   
         </div>
		</div>
	</div>
</body>
</html>
