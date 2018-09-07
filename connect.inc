 <%
  'Create connection string
  strDBName="DB_109590_dmsdev"
  strUID = "sa"
  strPwd="Sanchin1153!"
  strDS="LAPTOP-FFV1U9QC"
  
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
 
    ' create recordset e and set its activeconnection 
  set orse = server.createobject("adodb.recordset") 
  oRSE.ActiveConnection = oConn 
  
    ' create recordset f and set its activeconnection 
  set orsf = server.createobject("adodb.recordset") 
  oRSf.ActiveConnection = oConn 

      ' create recordset G and set its activeconnection 
  set orsG = server.createobject("adodb.recordset") 
  orsG.activeconnection = oconn 
  
       ' create recordset H and set its activeconnection 
  set orsH = server.createobject("adodb.recordset") 
  orsH.activeconnection = oconn 

       ' create recordset J and set its activeconnection 
  set orsJ = server.createobject("adodb.recordset") 
  orsJ.activeconnection = oconn 

  
      ' create recordset L and set its activeconnection 
  set orsL = server.createobject("adodb.recordset") 
  orsL.activeconnection = oconn 
  
      ' create recordset p and set its activeconnection 
  set orsp = server.createobject("adodb.recordset") 
  orsp.activeconnection = oconn 
  
    ' create recordset r and set its activeconnection 
  set orsr = server.createobject("adodb.recordset") 
  orsr.activeconnection = oconn 

      ' create recordset s and set its activeconnection 
  set orss = server.createobject("adodb.recordset") 
  orss.activeconnection = oconn 

   ' Create RecordSet T and set its ActiveConnection 
  Set oRST = Server.CreateObject("ADODB.Recordset") 
  oRST.ActiveConnection = oConn 
  
   ' Create RecordSet X and set its ActiveConnection 
  Set oRSX = Server.CreateObject("ADODB.Recordset") 
  oRSX.ActiveConnection = oConn 
 
  
   ' Create RecordSet Y and set its ActiveConnection 
  Set oRSY = Server.CreateObject("ADODB.Recordset") 
  oRSY.ActiveConnection = oConn 

  
  ' Create RecordSet Z and set its ActiveConnection 
  Set oRSZ = Server.CreateObject("ADODB.Recordset") 
  oRSZ.ActiveConnection = oConn 
 
    ' Create RecordSet X and set its ActiveConnection 
  Set oRSX = Server.CreateObject("ADODB.Recordset") 
  oRSX.ActiveConnection = oConn 

  %>                    