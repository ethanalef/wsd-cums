<%
session("dataConn")="sql"

if session("dataConn") = "access" then
	thisFileName=Request.ServerVariables("script_name")
	thisFileName=mid(thisFileName,InstrRev(thisFileName,"/")+1)
	if thisFileName="index.asp" then
		mdbPath= Server.MapPath("dbf2000.mdb")
	else
		mdbPath= Server.MapPath("..\dbf2000.mdb")
	end if
	strconn = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & mdbPath
else
	strconn= "Driver={SQL Server};Server=XP-NOTEBOOK1;Database=wsdscu"
end if

set conn = server.createobject("adodb.connection")
conn.open strconn
%>
