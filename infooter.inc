

<%
if (bitIsAdmin or bitIsSysAdmin) and bitLoggedIn then
%><a href="index.asp"><img src="../images/home.gif" border="1"  /></a>&nbsp;&nbsp;&nbsp;
<%
end if     
%>
<span style="text-align:right; color:navy; font-size:10px;"> User:  <b><%=session("Username")%></b>&nbsp;&nbsp; Last Access:  <b><%=Session("LastAccessDate")%></b> Kai <%=strKaiCd%> Locale <%=strLocale%> Stage <%=strStage%>
<%  
if session("IsAdmin") then 
   response.write(" Admin") 
end if
if session("IsSysAdmin") then 
   response.write(" SysAdmin")
end if
if session("GodMode") then
   response.write(" GM")
end if   
  %>
</span> 
&nbsp;&nbsp;<span style="color:darkgrey; font-size:10px;"><%=strStagestr%>&nbsp; Copyright 2016 New Paltz Karate Academy  </span>


      