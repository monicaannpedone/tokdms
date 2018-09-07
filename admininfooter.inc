&nbsp;
<a href=logout.asp ><img src="../images/logout.gif" border="1"  /></a>


<%
if (bitIsAdmin or bitIsSysAdmin) and bitLoggedIn then
%><a href="index.asp"><img src="../images/home.gif" border="1"  /></a>&nbsp;&nbsp;&nbsp;
<%
	end if     
%>
<span style="text-align:right; color:navy; font-size:10px;"> User:  <b><%=session("Username")%></b>&nbsp;&nbsp; Last Access:  <b><%=Session("LastAccessDate")%></b></span>
&nbsp;&nbsp;<span style="color:darkgrey; font-size:10px;"><%=strStagestr%> </span>  


      