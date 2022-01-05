<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

SQl = "select * from memMaster where deleted=0 and thisShrBal"&acPeriod&"<5 order by memNo"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
elseif request.form("output")="text" then
	spaces=""
	for idx = 1 to 50
		spaces=spaces&" "
	next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Server.MapPath("..\txt")&"\"&session("username")&".txt", True)
	objFile.Write "EMSD Credit Union"
	objFile.WriteLine ""
	objFile.Write "Dormant Account List"
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write right(spaces&"Share Balance",20)
	objFile.WriteLine ""
	for idx = 1 to 80
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write right(spaces&formatnumber(rs("thisShrBal"&acPeriod),2),20)
		objFile.WriteLine ""
		rs.movenext
	loop
	objFile.Close
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>Dormant Account List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="4">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">Dormant Account List</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="200"><b>Name</b></td>
		<td width="130" align="right"><b>Share Balance</b></td>
	</tr>
	<tr><td colspan=3><hr></td></tr>
<%
do while not rs.eof
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td align="right"><%=formatNumber(rs("thisShrBal"&acPeriod),2)%></td>
	</tr>
<%
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
