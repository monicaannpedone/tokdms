<%
		'Response.Expires = 0
		'if(Application("Down") and NOT Session("IsAdmin")) then
		'   response.redirect("SiteDown.asp")
		'end if
		'Ensure that this page is not cached.
		
		IF Session("loggedin") = false then
		     response.redirect("login.asp")
		ELSE
			Session("loggedin")=True
		END IF
%>