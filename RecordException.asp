<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validate.inc"-->
<!--#include file="../include/connect.inc"-->
<!-- #include file ="../include/locale.inc" -->

<HTML>
<head>
<title>Error Handler</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
    </head>
<BODY>
<h1>Error Handler</h1>
<h2> Error Handler</h2>
<%        strErrCode = Request.QueryString("ErrCode")
          strErrMsg  = Request.QueryString("ErrMsg")
		  strMod     = Request.QueryString("Mod")
		 
		  response.write("Code   >"&request.querystring("errcode")&"<")
		  response.write("Msg    >"&request.querystring("errmsg")&"<")
          response.write("Module >"&request.querystring("mod")&"<")  
  strSQL = "INSERT INTO dms_Incident (IncidentDateTime, IncidentModule,     IncidentCode,     IncidentText)"
  strSQL = strSQL &         " VALUES ('"&Now()&"',        '"&strMod&"', '"&strErrCode&"', '"&strErrMsg&"' );"
  response.write("ErrCode=>"&strErrCode&"< <br>")
  response.write("ErrMsg=>"&strErrMsg&"< <br>")
  response.write("Mod=>"&strMod&"< <br>")
  response.write("StrSQL=>"&strSQL&"< <br>")
  response.End()
  
  oConn.Execute(strSQL)
  
  set oRS = Nothing
  

%>		  
     
</BODY>
</HTML> 