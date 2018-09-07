 <%
  'Create connection string
  
  strDBName     = "mtsdev"
  strUID        = "sa"
  strPwd        = "Dark6day:"
  strDataSource = "IBM-0EFB26BA722\SQLEXPRESS"
  strProvider   = "SQLOLEDB"
 
  strConnectionString = "Provider="&strProvider&";Data Source="&strDataSource&";Initial Catalog="&strDBName
  strConnectionString = strConnectionstring&";User Id="&strUID&";Password="&strPWD&";"
 
  'Create Connection
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.ConnectionTimeout = 60
  oConn.Open strConnectionString
  
  Set OCmd = Server.CreateObject("ADODB.Command")
  Set OCmd.ActiveConnection = oConn
  
  ' Create RecordSet and set its ActiveConnection
  Set oRS = Server.CreateObject("ADODB.Recordset")
  oRS.ActiveConnection = oConn
  
 %>

