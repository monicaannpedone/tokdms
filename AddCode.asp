<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Add Code</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
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
	  <h2>Add a Code <%=strCodeName%></h2>
            <Form action="AddCode1.asp" method="post" name="AddCode1" >
                  <input type="hidden" name="CodeName" value="<%=strCodename%>" /> 
                  
            <table class="dtab">
            <%
            strSQL="SELECT Codename, CodeValue, CodeValueText, ValueOrder, ValueIsActive FROM dms_Codes WHERE CodeName='"&strCodeName&"' ORDER BY ValueOrder;"
            bitFlop = True
            icnt = 0
            oRS.Open strSQL
            %>
            <tr>
               <td class="dtabh">Value</td>
               <td class="dtabh">Text</td>
               <td class="dtabh">Order</td>
               <td class="dtabh">Active?</td>
               <td class="dtabh">&nbsp;</td>
            </td>   
            <%
            intOrderMax = 0
            do while NOT oRS.EOF
               inct = icnt + 1
               strCodeValue     = ors("CodeValue")
               strCodeValueText = ors("CodeValueText")
               intValueOrder    = ors("ValueOrder")
               bitValueIsActive = ors("ValueIsActive")
               if intValueOrder > intOrderMax then
                  intOrderMax = intValueOrder
               end if   
               %>
               <tr>
               <td><%=strCodeValue%></td> 
               <td><%=strCodeValueText%></td>
               <td><%=intValueOrder%></td> 
               <td><%=bitValueIsActive%></td>
               <td>&nbsp;</td>
               </tr>
            <%
			   ors.MoveNext
            loop
            intOrderMax = IntOrderMax + 1  
            %>
               <tr>                  
                  <td class="drowedit"  align="left"><input name="CodeValue" type="text"  class="entry" size = "3" length="3" maxlength="3" value="" /></td>
                  <td class="drowedit"><input name="CodeDesc" type="text" class="entry"   /></td>
                  <td class="drowedit"><input name="Valueorder" value="<%=intOrderMax%>" class="entry"/></td>
                  <td class="drowedit"><input type="checkbox" name="ValueIsActive" value="Y" CHECKED></td>
                  <td class="drowerror" ><%=strErrMsg%></td>
               </tr> 
               <tr>                  
                  <td class="drowname" align="left"><input  type="Reset" value="Reset"></td>
                  <td class="drownone" ></td>
                  <td class="drownone" ></td>
                  <td class="drownone" >&nbsp;</td>
                  <td class="drowname"  align="right"><input  type="submit" value="Add"></td>
               </tr> 
			</table>
            </Form>
            <br>
            <a href="Codemaint.asp">Done - Return to Code List</a>
        </div> 
    <!--#include file="../include/infooter.inc"-->     
        
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
