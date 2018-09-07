<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validate.inc"-->
<!--#include file="../include/connect.inc"-->
<!--#include file="../include/locale.inc"-->
<!--#include file="../include/common.inc"-->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Display Codes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/regular.js"></script>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
</head>

<body>
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	strCodename = Request.QueryString("cd")
	strCodeValue = Request.QueryString("cv")
	strErrMsg    = Request.QueryString("ErrMsg")
      
	%>

	<div id="main">
	    <% Call TopLine() %>
		<div id="contents">
			<h2>Display Codes <%=strTitle%></h2>
            <h3></h3>
            <table>
            <div class="drowheader">
            <tr>
               <th class="dtabh">CodeName</th>
               <th class="dtabh">Order</th>
               <th class="dtabh">Value</th>
               <th class="dtabh">Value Text</th>
               <th class="dtabh">Is Active</th>
            </tr>
            </div>   
            <%			
			strSQL="SELECT CodeName, ValueOrder, CodeValue, CodeValuetext, ValueIsActive FROM dms_Codes ORDER BY CodeName, ValueOrder;"
            
			bitFlop = True
            icnt = 0 
            strOldCodename = ""
            intCodeCount = 0
            oRS.Open strSQL
            bitParity =1
            do while NOT ors.eof
               icnt = icnt + 1
               
               strCodeName          = ors("CodeName")
               strValueOrder        = ors("ValueOrder")
               strCodeValue         = ors("CodeValue")
               strCodeValueText     = ors("CodeValueText")
               bitValueIsActive     = ors("ValueIsActive")
                if strOldCodeName = strCodeName then
                else
                   intCodeCount = intCodeCount +1
                   strOldCodeName = strCodeName
                end if   
               if bitValueIsActive then
                  strValueIsActive = "Active"
               end if
               if bitParity = 1 then
                  strClass="drowodd"
               else
                  strClass="droweven"
               end if      
               bitParity = NOT bitParity
            %>
               <tr class="<%=strClass%>"><td><%Response.write(strCodename)%></td>
                   <td><%response.write(strValueOrder)%></td>
                   <td><%response.write(strCodeValue)%></td>
                   <td><%response.write(strCodeValueText)%></td>
                   <td><%response.write(strValueIsActive)%></td>
               </tr>
			<%
               oRS.MoveNext 
            loop
            
            if icnt = 0 then
			      strMsg ="The Code Table has no entries" 
                  response.end
                  response.redirect("Errorhandler.asp?ErrCode=NOTFound&DisplayStatusCodes&Msg="&strMsg)
            end if   

			%> 
            <tr>
              <td class="dtabh" colspan="5">*** Number of Codes: <%=intCodeCount%>; Items: <%=icnt%> ***</td>
            </tr>  
            </table>
            <%  
	   ors.Close
       set ors = Nothing
	   %> 
        </div> 
	</div>
</body>
</html>
