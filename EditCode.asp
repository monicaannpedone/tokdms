<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Add Physical Media Code</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/regular.js"></script>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>

</head>

<body>
<div id="main">
    <div id="titlecontents">
        <div id="TitleBox">
            <h2><a href="Admin.asp"><i class="fa fa-bars"></i></a>&nbsp;&nbsp;<img src="../images/lillilinvertedtomoe.gif" />&nbsp;Dojo Management System </h2>
        </div>   
        <div id="AcctBox">
            <h4><a href="admin.asp"><%=strUserName%>&nbsp;<i class="fa fa-user"></i></a>&nbsp;&nbsp;&nbsp;<a href="sysadmin.asp"><i class="fa fa-cog"></i></a>&nbsp;&nbsp;<a href="Logout.asp"><i class="fa fa-sign-out"></i></a></h4>
        </div>
    </div>  
	  <div id="contents">
	  <div id="OneCol">
	  <h1>Edit Code</h1>
	  <h3></h3>
      <a href="CodeMaint.asp">Return to Code Maintenance</a>
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	strCodename = Request.QueryString("cd")
	strErrMsg    = Request.QueryString("ErrMsg")
	'response.write("strCodeName=>"&strCodeName&"<<br>")
	'response.write("strErrMsg  =>"&strErrMsg&"<<br>")
	
	if len(trim(strCodeName)) = 0 then
	  response.redirect("AddCodeName.asp")
    end if 
	%>
	  <br><br>
	  <a href="AddCode.asp?cd=<%=strCodeName%>"><i class="far fa-plus"></i>&nbsp;&nbsp;Add a Code Value to Code <%=strCodename%></a></br>
          
                 
                  
            <table class="dtab">
            <%
            strSQL="SELECT Codename, CodeValue, CodeValueText, ValueOrder, ValueIsActive FROM dms_Codes WHERE CodeName='"&strCodeName&"' ORDER BY ValueOrder;"
            bitFlop = True
            icnt = 0
            oRS.Open strSQL
            %>
            <tr>
               <td class="dtabnone">&nbsp;</td>
               <td class="dtabh">Value</td>
               <td class="dtabh">Text</td>
               <td class="dtabh">Order</td>
               <td class="dtabh">Active?</td>
               <td class="dtabnone">&nbsp;</td>
            </td>   
            <%
            intOrderMax = 0
            bitFlop = 1
            do while NOT oRS.EOF
               inct = icnt + 1
               if bitFlop = 1 then
                  strClass="drowodd"
               else
                  strClass="droweven"
               end if      
               strCodeValue     = ors("CodeValue")
               strCodeValueText = ors("CodeValueText")
               intValueOrder    = ors("ValueOrder")
               bitValueIsActive = ors("ValueIsActive")
           %>
               <tr class="<%=strClass%>">
               <td class="dtabedit" align="center" ><a href="EditCode0.asp?cd=<%=strCodename%>&cv=<%=strCodeValue%>" > <i class="far fa-edit"></i></a></td> 
               <td align="center" ><%=strCodeValue%></td> 
               <td><%=strCodeValueText%></td>
               <td align="center" ><%=intValueOrder%></td> 
               <td align="center" ><%=bitValueIsActive%></td>
               <td class="dtabedit" align="center" ><a href="deleteCode0.asp?cd=<%=strCodename%>&cv=<%=strCodeValue%>" > <i class="far fa-trash-alt"></i></a></td> 
               </tr>
            <%
               bitFlop  = NOT bitFlop
			   ors.MoveNext
            loop
            %>
            <tr>
               <td class="dtabnone">&nbsp;</td>
               <td class="dtabh" colspan="4">&nbsp;</td>
               <td class="dtabnone">&nbsp;</td>
            </td>   
			</table>
        </div> 
      </div>  
	</div>
</body>
</html>
