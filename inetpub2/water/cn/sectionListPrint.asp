<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

f = ""
a = split(request("TS"),",",-1,1)
if ubound(a) < 0 then
	response.redirect "sectionList.asp"
end if

for i = 0 to ubound(a)
	f = f & "'"&trim(A(i))&"',"
next
f = left(f,len(f)-1)

sql = "select * from memMaster where deleted=0 and  memSection In ("&f&") order by memSection,memNo;"
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
	objFile.Write "Section List"
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write right(spaces&"Share B/L",20)
	objFile.Write right(spaces&"Loan B/L",20)
	objFile.Write right(spaces&"SD",20)
	objFile.Write right(spaces&"Autopay",20)
	objFile.WriteLine ""
	for idx = 1 to 140
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		if thisSection<>rs("memSection") then
			thisSection=rs("memSection")
			objFile.Write rs("memSection")
			objFile.WriteLine ""
		end if
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write left(formatnumber(rs("thisShrBal"&acPeriod),2),20)
		objFile.Write left(formatnumber(rs("thisLoanBal"&acPeriod),2),20)
		objFile.Write left(formatnumber(rs("OSInterest"),2),20)
		objFile.Write left(formatnumber(rs("salaryDedut"),2),20)
		objFile.Write left(formatnumber(rs("autopayAmt"),2),20)
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
<title>Section List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="7"><font size="5">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="7"><font size="5">Section List</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="200"><b>Name</b></td>
		<td width="100" align="right"><b>Share B/L</b></td>
		<td width="100" align="right"><b>Loan B/L</b></td>
		<td width="100" align="right"><b>Interest</b></td>
		<td width="100" align="right"><b>Salary<br>Deduction</b></td>
		<td width="100" align="right"><b>Autopay</b></td>
	</tr>
	<tr><td colspan=7><hr></td></tr>
<%
thisSection=""
do while not rs.eof
	if thisSection<>rs("memSection") then
		thisSection=rs("memSection")
%>
	<tr>
		<td colspan="3" height="25" valign="middle"><b>Section : <%=rs("memSection")%></b></td>
	</tr>
<%
	end if
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td align=right><%=formatNumber(rs("thisShrBal"&acPeriod),2)%></td>
		<td align=right><%=formatNumber(rs("thisLoanBal"&acPeriod),2)%></td>
		<td align=right><%=formatNumber(rs("OSInterest"),2)%></td>
		<td align=right><%=formatNumber(rs("salaryDedut"),2)%></td>
		<td align=right><%=formatNumber(rs("autopayAmt"),2)%></td>
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
