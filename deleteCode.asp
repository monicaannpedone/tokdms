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
	<style type="text/css" media="screen">@import "../css/font-awesome.min.css";</style>  
	<script src="https://use.fontawesome.com/2ad953023c.js"></script> 

</head>

<body>
	
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
	<div id="header">
	
  
	</div>
	<div id="main">
		<div id="contents">
	  <h2>Edit Code values for Code <%=strCodeName%></h2>
	  <a href="AddCode.asp?cd=<%=strCodeName%>"><i class="fa fa-plus"></i>&nbsp;&nbsp;Add a Code Value to Code <%=strCodename%></a></br>
      </br>      
                 
                  
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
            
            do while NOT oRS.EOF
               inct = icnt + 1
               strCodeValue     = ors("CodeValue")
               strCodeValueText = ors("CodeValueText")
               intValueOrder    = ors("ValueOrder")
               bitValueIsActive = ors("ValueIsActive")
           %>
               <tr>
               <td class="dtabedit" align="center" ><a href="EditCode0.asp?cd=<%=strCodename%>&cv=<%=strCodeValue%>" > <i class="fa fa-edit"></i></a></td> 
               <td align="center" ><%=strCodeValue%></td> 
               <td><%=strCodeValueText%></td>
               <td align="center" ><%=intValueOrder%></td> 
               <td align="center" ><%=bitValueIsActive%></td>
               <td class="dtabedit" align="center" ><a href="deleteCode0.asp?cd=<%=strCodename%>&cv=<%=strCodeValue%>" > <i class="fa fa-remove"></i></a></td> 
               </tr>
            <%
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
    <!--#include file="../include/infooter.inc"-->     
        
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
