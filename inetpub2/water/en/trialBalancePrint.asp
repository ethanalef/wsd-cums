<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

mMonth = int(left(request("mPeriod"),2))
mYear = int(right(request("mPeriod"),4))

if mMonth<=4 then
	m=mMonth+8
	if mYear=acYear then y=mYear else y=mYear-1 end if
else
	m=mMonth-4
	if mYear=acYear then y=mYear+1 else y=mYear end if
end if
mDate = y&"/"&m&"/"&day(dateAdd("m",1,y&"/"&m&"/1")-1)

if mYear=acYear then y="this" else y="last" end if

SQl = "select glId,glName,"&y&"Bal"&int(mMonth)&" as bal from glMaster where deleted=0 and creationDate<='"&mDate&"' order by glId"
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
	objFile.Write "WSDS Credit Union"
	objFile.WriteLine ""
	objFile.Write "Trial Balance"
	objFile.WriteLine ""
	objFile.Write "Account Year : "&mYear&"  Period : "&mMonth
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("A/C Title"&spaces,50)
	objFile.Write right(spaces&"Debit",20)
	objFile.Write right(spaces&"Credit",20)
	objFile.WriteLine ""
	for idx = 1 to 100
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		objFile.Write left(rs("glId")&spaces,10)
		objFile.Write left(rs("glName")&spaces,50)
		if rs("bal")>0 then
			objFile.Write right(spaces&formatnumber(rs("bal"),2),20)
			totaldb = totaldb+rs("bal")
		else
			objFile.Write left(spaces,20)
		end if
		if rs("bal")<=0 then
			objFile.Write right(spaces&formatnumber(-1*rs("bal"),2),20)
			totalcr = totalcr+rs("bal")
		else
			objFile.Write left(spaces,20)
		end if
		objFile.WriteLine ""
		rs.movenext
	loop
	for idx = 1 to 100
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write left(spaces&spaces,60)
	objFile.Write right(spaces&formatnumber(totaldb,2),20)
	objFile.Write right(spaces&formatnumber(-totalcr,2),20)
	objFile.WriteLine ""
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
<title>Trial Balance</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="4"><font size="4">WSDS Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="4"><font size="4">Trial Balance</font></td>
	</tr>
	<tr height="30" valign="top" align="center">
		<td colspan="4">Account Year : <%=mYear%> &nbsp; &nbsp; Period : <%=mMonth%></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="300"><b>A/C Title</b></td>
		<td width="130" align=right><b>Debit</b></td>
		<td width="130" align=right><b>Credit</b></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>
<%
	do while not rs.eof %>
	<tr>
		<td><%=rs("glId")%></td>
		<td><%=rs("glName")%></td>
<%
		if rs("bal")>0 then
			response.write "<td align=right>" & formatnumber(rs("bal"),2) & "</td><td></td>"
			totaldb = totaldb+rs("bal")
		else
			response.write "<td></td><td align=right>" & formatnumber(-1*rs("bal"),2) & "</td>"
			totalcr = totalcr+rs("bal")
		end if
		rs.movenext
	loop
%>
	<tr>
		<td colspan="2"></td>
		<td align="right"><hr></td>
		<td align="right"><hr></td>
	</tr>
	<tr>
		<td colspan="2" class="b10" align="right">Total : </td>
		<td align="right"><%=formatnumber(totaldb,2)%></td>
		<td align="right"><%=formatnumber(-totalcr,2)%></td>
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