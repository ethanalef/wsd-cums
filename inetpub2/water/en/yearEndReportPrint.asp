<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

mDividend=request("mDividend")
mYear=request("year")

if mYear="this" then
	conn.execute("update memMaster set ttlShare=thisShrBal1 where deleted=0")
	for idx = 2 to 12
		conn.execute("update memMaster set ttlShare=ttlShare + FLOOR(thisShrBal"&idx&" / 5) * 5 where deleted=0")
		conn.execute("update memMaster set ttlShare=FLOOR(thisShrBal"&idx&" / 5) * 5 * "&idx&" where thisShrBal"&idx&"<thisShrBal"&idx-1&" and deleted=0")
	next

	sql = "select memNo,memName,round(ttlShare/12*"&mDividend&"/100,2) as dividend from memMaster where deleted=0 order by memNo"
else
	conn.execute("update memMaster set ttlLastShare=lastShrBal1 where deleted=0")
	for idx = 2 to 12
		conn.execute("update memMaster set ttlLastShare=ttlLastShare + FLOOR(lastShrBal"&idx&" / 5) * 5 where deleted=0")
		conn.execute("update memMaster set ttlLastShare=FLOOR(lastShrBal"&idx&" / 5) * 5 * "&idx&" where lastShrBal"&idx&"<lastShrBal"&idx-1&" and deleted=0")
	next

	sql = "select memNo,memName,round(ttlLastShare/12*"&mDividend&"/100,2) as dividend from memMaster where deleted=0 order by memNo"
end if

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
	objFile.Write "Year End Report"
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write right(spaces&"Dividend",15)
	objFile.WriteLine ""
	for idx = 1 to 75
		objFile.Write "-"
	next
	objFile.WriteLine ""
	ttlAmt=0
	do while not rs.eof
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write right(spaces&formatnumber(rs("dividend"),2),15)
		objFile.WriteLine ""
		ttlamt=ttlamt+round(rs("dividend"),2)
		rs.movenext
	loop
	for idx = 1 to 75
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"Total : ",60)
	objFile.Write right(spaces&formatnumber(ttlamt,2),15)
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
<title>Year End Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="4">WSDS Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">Year End Report</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="200"><b>Name</b></td>
		<td width="130" align="right"><b>Dividend</b></td>
	</tr>
	<tr><td colspan=3><hr></td></tr>
<%
ttlAmt=0
do while not rs.eof
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td align="right"><%=formatNumber(rs("dividend"),2)%></td>
	</tr>
<%
	ttlamt=ttlamt+round(rs("dividend"),2)
	rs.movenext
loop
%>
	<tr><td colspan=3><hr></td></tr>
	<tr>
		<td colspan="2" class="b10" align="right">Total : </td>
		<td align="right"><%=formatNumber(ttlamt,2)%></td>
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
