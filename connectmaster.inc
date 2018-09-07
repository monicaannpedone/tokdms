 <%
  'Create connection string
  strDBName="master"
  strUID = "sa"
  strPwd="Sanchin1153!"
  strDS="MONICAANNPE8315\SQLEXPRESS"
  
  strsource="Provider=sqloledb;uid="&strUID&";pwd="&strPwd&";Data Source="&strDS&"; Initial Catalog ="&strDBName&";"
  strConnectionString = strsource

  	'Create Connection
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.ConnectionTimeout = 60
  oConn.Open strConnectionString

  ' Create RecordSet and set its ActiveConnection 
  Set oRS = Server.CreateObject("ADODB.Recordset")
  oRS.ActiveConnection = oConn
  
 

  ' Create RecordSet A and set its ActiveConnection 
  Set oRSA = Server.CreateObject("ADODB.Recordset")
  oRSA.ActiveConnection = oConn 
  
  ' Create RecordSet B and set its ActiveConnection 
  Set oRSB = Server.CreateObject("ADODB.Recordset")
  oRSB.ActiveConnection = oConn 

  ' Create RecordSet C and set its ActiveConnection 
  Set oRSC = Server.CreateObject("ADODB.Recordset") 
  oRSC.ActiveConnection = oConn                     

  ' Create RecordSet D and set its ActiveConnection 
  Set oRSD = Server.CreateObject("ADODB.Recordset") 
  oRSD.ActiveConnection = oConn 
  
   ' Create RecordSet T and set its ActiveConnection 
  Set oRST = Server.CreateObject("ADODB.Recordset") 
  oRST.ActiveConnection = oConn 
  
 
  
  ' Create RecordSet Z and set its ActiveConnection 
  Set oRSZ = Server.CreateObject("ADODB.Recordset") 
  oRSZ.ActiveConnection = oConn 
  
    ' Create RecordSet X and set its ActiveConnection 
  Set oRSX = Server.CreateObject("ADODB.Recordset") 
  oRSX.ActiveConnection = oConn 

  %>                    