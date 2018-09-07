<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Change Dojo Code</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>
<div >
   <div >
	  <br/><br/>
		
<% 
strModuleNo = Request.QueryString("mn")
response.write("MN=>"&strModuleNo&"<<br>")
intfunc = Request.Querystring("f")
if strModuleNo = "" then
   strModule = "MonthViewAttendance.asp"
end if
if strModuleNo = "1" then
   strModule = "ListStudent.asp"
end if
if strModuleNo = "2" then
   strModule = "AssignGender.asp"
end if
if strModuleNo = "3" then
   strModule = "StudentFunction.asp?f="&intFunc
end if
   
strDojoCd = Trim(Request.Form("DojoCd"))
intLenDojoCd = Len(strDojoCd)
strSelActive= Request.Form("selactive")
strSelInactive = request.form("selinactive")
strUserName=Session("UserName")
response.write("UserName=>"&strUserName&"< </br>")
response.write("Dojocd=>"&strDojocd&"< </br>")
response.write("SelActive=>"&strSelActive&"< </br>")
response.write("SelInActive=>"&strSelInActive&"< </br>")
strSetActive = 1
strSetInactive = 0
if isNull(strSelActive) then
   strSetActive=0
else
   if strSelActive = "active" then
      strSetActive = 1
   else 
      strSetActive = 0
   end if      
end if   
if isNull(strSelInActive) then
   strSetInActive=0
else
   if strSelInActive = "inactive" then
      strSetInActive = 1
   else 
      strSetInActive = 0
   end if      
end if   
response.write("UserName=>"&strUserName&"< </br>")
response.write("Dojocd=>"&strDojocd&"< </br>")
response.write("SelActive=>"&strSelActive&"< </br>")
response.write("SelInActive=>"&strSelInActive&"< </br>")
if intlendojoCd = 0 then
   strChangeDC = ""
else 
   strChangeDC =  "DojoCd = '"&strDojoCD&"', "
end if    
'strSQL = "UPDATE  dms_pref SET DojoCd = '"&strDojoCD&"', SelActive = "&strSetActive&", SelInactive="&StrSetInactive&" WHERE UserName = '"&strUserName&"';"
strSQL = "UPDATE  dms_pref SET "&strChangeDC&" SelActive = "&strSetActive&", SelInactive="&StrSetInactive&" WHERE UserName = '"&strUserName&"';"

response.write("strSQL=>"&strSQL&"<</br>")

oConn.Execute(strSQL)

response.redirect(strModule)


 
%> 
		</div>
        <!--#include file="../include/infooter.inc"-->  
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
