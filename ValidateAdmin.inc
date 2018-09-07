<%
		Response.Expires = 0
		if(Application("Down") and NOT Session("IsAdmin")) then
		   response.redirect("ForceEnd.asp")
		end if
		'Ensure that this page is not cached.
		
		IF Session("loggedin") = false or Session("IsAdmin") = false then
		    response.redirect("login.asp")
		ELSE
			Session("loggedin")=True
		END IF
%>