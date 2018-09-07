<%
strAccessPath = "/dms/dev/access/Database11.accdb"

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
Dim connString
connString = "Provider=Microsoft.ACE.OLEDB.12.0; "  & "Data Source=" & Server.MapPath(strAccessPath&";")
conn.Open connString
Set rsAccess = Server.CreateObject("ADODB.Recordset")
rsAccess.ActiveConnection = conn
Set rsAccess1 = Server.CreateObject("ADODB.Recordset")
rsAccess1.ActiveConnection = conn
Set rsAccess2 = Server.CreateObject("ADODB.Recordset")
rsAccess2.ActiveConnection = conn
Set rsAccess3 = Server.CreateObject("ADODB.Recordset")
rsAccess3.ActiveConnection = conn
Set rsAccess4 = Server.CreateObject("ADODB.Recordset")
rsAccess4.ActiveConnection = conn
Set rsAccess5 = Server.CreateObject("ADODB.Recordset")
rsAccess5.ActiveConnection = conn
%>