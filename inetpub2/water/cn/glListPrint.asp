<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
dim mType(6)
mType(1) = "Fixed Assets"
mType(2) = "Loans"
mType(3) = "Current Assets"
mType(4) = "Liabilities"
mType(5) = "Income"
mType(6) = "Expenses"

SQl = "select glId,glName,glType from glMaster where deleted=0 order by glType,glId"
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
	objFile.Write "G/L List"
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("A/C Title"&spaces,50)
	objFile.Write left("Category"&spaces,20)
	objFile.WriteLine ""
	for idx = 1 to 80
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		objFile.Write left(rs("glId")&spaces,10)
		objFile.Write left(rs("glName")&spaces,50)
		objFile.Write left(mType(rs("glType"))&spaces,20)
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
<title>G/L List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="4"><font size="4">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="4"><font size="4">G/L List</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="350"><b>A/C Title</b></td>
		<td width="100"><b>Category</b></td>
	</tr>
	<tr><td colspan=3><hr></td></tr>
<%
do while not rs.eof %>
	<tr>
		<td><%=rs("glId")%></td>
		<td><%=rs("glName")%></td>
		<td><%=mType(rs("glType"))%></td>
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