<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include file="../include/locale.inc"-->

<HTML>
<HEAD>
<TITLE>Main Menu</TITLE>
</HEAD>
<BODY>
<%
    if (Session("loggedin")) then 
       response.Write("<br>"&Session("UserName")&" has been successfully logged out of tcs.<br><br>")
    else
       response.Write("<br> There is no active session. Session ended unconditionally<br><br>")
    end if
	Session.Abandon
%>

</BODY>
</HTML>