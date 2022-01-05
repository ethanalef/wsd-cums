<%
Set conn = Server.CreateObject("adodb.connection")
conn.Open "Driver={SQL Server};Server=WSDSCU-SVR;Database=wsdscu"
cuname=""
FMonthStart = 8
ArrMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
SETLOCALE(1033)
%>
