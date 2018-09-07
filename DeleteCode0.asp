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
</head>

<body>
	
	<%
	bitIsAdmin    = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn   = Session("loggedin")
	strCodename  = Request.QueryString("cd")
	strCodeValue = Request.QueryString("cv")
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
	  <h2>Confirm Delete Value <%=strCodeValue %> for Code <%=strCodeName%> </h2>
            <Form action="DeleteCode1.asp" method="post" name="DeleteCode1" >
                  <input type="hidden" name="CodeName" value="<%=strCodename%>" /> 
                  <input type="hidden" name="CodeValue" value="<%=strCodeValue%>" /> 

            <table class="dtab">
            <%
            strSQL="SELECT Codename, CodeValue, CodeValueText, ValueOrder, ValueIsActive FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND CodeValue='"&strCodeValue&"';"
            bitFlop = True
            icnt = 0
            oRS.Open strSQL
            %>
            <tr>
               <td class="dtabh">Name</td>
               <td class="dtabh">Value</td>
               <td class="dtabh"width="100">Text</td>
               <td class="dtabh">Order</td>
               <td class="dtabh">Active?</td>
            </td>   
            <%
            intOrderMax = 0
            if oRS.EOF then
               response.write("ERROR, cant find Code Name >"&strCodeName&"<, Code Value = >"&strCodeValue&"<<br>")
               response.end
            end if
            strCodeValue     = ors("CodeValue")
            strCodeValueText = ors("CodeValueText")
            intValueOrder    = ors("ValueOrder")
            bitValueIsActive = ors("ValueIsActive")
            if bitValueIsActive then
               strChecked = "Checked"
            else
               strChecked = ""
            end if      
            %>
               <tr>  
                  <td><%=strCodeName%></td>                
                  <td align="center"><%=strCodeValue%></td>
                  <td ><%=strCodeValueText%></td>
                  <td align="center"><%=intValueOrder%></td>
                  <td align="center"><%=bitValueIsActive%></td>
               </tr>
               <tr>                  
                  <td class="drowname" align="left"><input  type="Reset" value="Reset"></td>
                  <td class="drowname" colspan="4"  align="right"><input  type="submit" value="Delete"></td>
               </tr> 
			</table>
            </Form>
        </div> 
    <!--#include file="../include/infooter.inc"-->     
        
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
