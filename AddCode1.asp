<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Edit Code  - Processor</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
	
	<%
		strCodename      = Request.Form("CodeName")
        strCodeValue     = Request.Form("CodeValue")
        strCodeDesc      = Request.Form("CodeDesc")
        intValueorder    = Request.Form("ValueOrder")
        bitValueIsActive = Request.Form("ValueIsActive")
        Response.write("strCodename=>"&strCodename&"<<br>")

        Response.write("Request.Form(CodeName)=>"&Request.Form("CodeName")&"<<br>")
        Response.write("strCodeValue=>"&strCodeValue&"<<br>")
        Response.write("strCodeDesc=>"&strCodeDesc&"<<br>")
        Response.write("intValueorder=>"&intValueorder&"<<br>")
        Response.write("bitValueIsActive=>"&bitValueIsActive&"<<br>")
        
        if bitValueIsActive = "Y" then
           intValueIsActive = 1
        else
           intValueIsActive = 0
        end if
              
    ' validate fields    
    if len(Trim(strCodeValue)) <1 then
       strEM = "The Code Value must be 1 to 3 Characters; Can not be all blank"
       response.redirect("AddCode.asp?cd="&strCodeName&"&ErrMsg="&strEM)
    end if
    if len(Trim(strCodeValue)) <1 then
       strEM = "The Code Value Text must not be all blank"
       response.redirect("AddCode.asp?cd="&strCodeName&"&ErrMsg="&strEM)
    end if
    if not IsNumeric(ValueOrder)then
       strEM = "The Code Order  must be a Numeric Integer"
       response.redirect("AddCode.asp?cd="&strCodeName&"&ErrMsg="&strEM)
    end if
    strSQL = "SELECT CodeName, CodeValue FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND  CodeValue='"&strCodeValue&"'; "
    oRS.Open strSQL
    if NOT oRS.EOF then
	   strEM = "That CodeName and Value, >"&strCodeName&"<, >"&strCodeValue&"< already exists; "
	   response.redirect("AddCode.asp?cd="&strCodename&"&ErrMsg="&strEM)
	end if
	  
	strSQLI = "INSERT INTO dms_codes (CodeName, CodeValue, CodeValueText, ValueOrder, ValueIsActive )"
	strSQLI = strSQLI & " VALUES ('"&strCodeName&"', '"&strCodeValue&"', '"& StrCodeDesc&"', "&intValueOrder&","&intValueIsActive&" );" 	
	response.write("strSQLI=>"&strSQLI&"<<br>")
	oConn.Execute(strSQLI)
			
	ors.Close
    set ors = Nothing
	response.redirect("AddCode.asp?cd="&strCodeName)
	%> 
</body>
</html>
