<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
mMonth = int(left(request("mPeriod"),2))
mYear = int(right(request("mPeriod"),4))

SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

if mMonth<=4 then
	m=mMonth+8
	if mYear=acYear then y=mYear else y=mYear-1 end if
else
	m=mMonth-4
	if mYear=acYear then y=mYear+1 else y=mYear end if
end if
mDate = y&"/"&m&"/"&day(dateAdd("m",1,y&"/"&m&"/1")-1)

if mYear=acYear then y="this" else y="last" end if

SQl = "select glId,glName,glType,"&y&"Bal"&int(mMonth)&" as bal from glMaster where deleted=0 and glType>=5 and creationDate<='"&mDate&"' order by glType,glId"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3

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
	objFile.Write "Profit & Loss Statement"
	objFile.WriteLine ""
	objFile.Write "Account Year : "&mYear&"  Period : "&mMonth
	objFile.WriteLine ""
	objFile.WriteLine ""
	objFile.Write "Income :"
	objFile.WriteLine ""
	objFile.WriteLine ""
	income=0
	expenses=0
	do while not rs.eof
		if rs("glType")<>5 then
			exit do
		end if
		objFile.Write left(rs("glName")&spaces,50)
		objFile.Write right(spaces&formatnumber(-1*rs("bal"),2),20)
		objFile.WriteLine ""
		income = income + round(-1*rs("bal"),2)
		rs.movenext
	loop
	objFile.WriteLine ""
	objFile.Write left("Income Total :"&space(70),70)
	objFile.Write right(spaces&formatnumber(income,2),20)
	objFile.WriteLine ""
	objFile.WriteLine ""
	objFile.Write "Expenses :"
	objFile.WriteLine ""
	objFile.WriteLine ""
	do while not rs.eof
		if rs("glType")<>6 then
			exit do
		end if
		objFile.Write left(rs("glName")&spaces,50)
		objFile.Write right(spaces&formatnumber(rs("bal"),2),20)
		objFile.WriteLine ""
		expenses = expenses + round(rs("bal"),2)
		rs.movenext
	loop
	objFile.WriteLine ""
	objFile.Write left("Expenses Total :"&spaces,50)
	objFile.Write right(spaces&formatnumber(expenses,2),20)
	objFile.WriteLine ""
	objFile.WriteLine ""
	objFile.Write left("Net Profit & Loss : "&spaces,70)
	objFile.Write right(spaces&formatnumber(income-expenses,2),20)
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
<title>Profit & Loss Statement</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="3"><font size="4">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="3"><font size="4">Profit & Loss Statement</font></td>
	</tr>
	<tr height="30" valign="top" align="center">
		<td colspan="3">Account Year : <%=mYear%> &nbsp; &nbsp; Period : <%=mMonth%></td>
	</tr>
	<tr height="35" valign="top">
		<td colspan="3"><font size="4">Income :</font></td>
	</tr>
<%
income=0
expenses=0
do while not rs.eof
	if rs("glType")<>5 then
		exit do
	end if
%>
	<tr>
		<td width="300"><%=rs("glName")%></td>
		<td width="130" align=right><%=formatnumber(-1*rs("bal"),2)%></td>
		<td width="130"></td>
	</tr>
<%
	income = income + round(-1*rs("bal"),2)
	rs.movenext
loop
%>
	<tr height="35" valign="top">
		<td><font size="4">Income Total :</font></td>
		<td></td>
		<td align=right><%=formatnumber(income,2)%></td>
	</tr>
	<tr height="35" valign="top">
		<td colspan="3"><font size="4">Expenses :</font></td>
	</tr>
<%
do while not rs.eof
	if rs("glType")<>6 then
		exit do
	end if
%>
	<tr>
		<td><%=rs("glName")%></td>
		<td align=right><%=formatnumber(rs("bal"),2)%></td>
		<td></td>
	</tr>
<%
	expenses = expenses + round(rs("bal"),2)
	rs.movenext
loop
%>
	<tr height="35" valign="top">
		<td><font size="4">Expenses Total :</font></td>
		<td></td>
		<td align=right><%=formatnumber(expenses,2)%></td>
	</tr>
	<tr height="35" valign="top">
		<td><font size="4">Net Profit & Loss :</font></td>
		<td></td>
		<td align=right><%=formatnumber(income-expenses,2)%></td>
	</tr>
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
