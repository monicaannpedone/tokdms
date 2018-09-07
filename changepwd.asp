<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/errormessage.inc" -->

<head>
	<title>Maintain Users</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>
<%
   strUserName = Session("Username")
%>   
<div >
   <div >
   <h2>Change Password for user <%=strUserName%></h2>
      <table>
         <form action="ProcChgPassword.asp" method="post" >
         <input type="hidden" name="user" value="<%=strUserName%>">
         <tr>
            <td class="dthead">Old Password</td><td class="dtedit"><input type="Text" name="OldPassword" value=""></td><td class="dcolerror"><%=aryErrMsg(01)%></td>
         </tr>
                  <tr>
            <td class="dthead">New Password</td><td class="dtedit"><input type="Text" name="NewPassword" value=""></td><td class="dcolerror"><%=aryErrMsg(02)%></td>
         </tr>
                  <tr>
            <td class="dthead">Confirm New Password</td><td class="dtedit"><input type="Text" name="ConfirmPassword" value=""></td><td class="dcolerror"><%=aryErrMsg(03)%></td>
         </tr>
                  <tr>
                <td  class="drowerror" colspan="3" align="left">
                   &nbsp;<%=aryerrMsg(04)%>
                </td>
         </tr> 

         <tr>
                <td  align="left">
                   <input type="Reset" value="Reset"  > 
                </td>
                   <td>&nbsp;</td>
                <td  align="right">
                   <input type="submit" value="Submit"  > 
                </td>
         </tr> 
        </form>

      </table>
      
   </div>
   <!--#include file="../include/infooter.inc"-->  
</div>
<!--#include file="../include/footer.inc"-->
</body>
</html>
