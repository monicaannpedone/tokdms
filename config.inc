<%
Function GetConfigValue(strName)
    strName=UCase(strName)
	strSQL = "SELECT cfg_value from tcs_Config WHERE cfg_name = '"&strName&"';"
	Set oConfig = Server.CreateObject("ADODB.Recordset")
	oConfig.ActiveConnection = oConn
	oConfig.Open strSQL
	if oConfig.EOF then
	   GetConfigValue = "*** "&strname&" NOT Found ***"
	else
	   GetConfigValue = trim(oConfig("cfg_value"))
	end if   
	oConfig.Close
	set oConfig = Nothing   
End Function
%>

