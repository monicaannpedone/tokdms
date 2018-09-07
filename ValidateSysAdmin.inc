<%
		Response.Expires = 0
		if(Application("Down") and NOT Session("IsSysAdmin")) then
		   response.redirect("ForceEnd.asp")
		end if
		'Ensure that this page is not cached.
		
		IF Session("loggedin") = false or Session("IsSysAdmin") = false then
		    response.redirect("login.asp")
		ELSE
			Session("loggedin")=True
		END IF
%>