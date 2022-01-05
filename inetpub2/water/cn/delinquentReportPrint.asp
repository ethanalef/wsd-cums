<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

SQl = "select * from memMaster where deleted=0 and thisShrBal"&acPeriod&">0 or thisLoanBal"&acPeriod&">0  order by memNo"
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
	objFile.Write "Delinquent Report"
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write right(spaces&"Loan Amount",20)
	objFile.Write right(spaces&"Over-due",20)
	objFile.Write right(spaces&"Share Amount",20)
	objFile.WriteLine ""
	for idx = 1 to 120
		objFile.Write "-"
	next
	objFile.WriteLine ""
	over1to2=0: amt1to2=0: over3to6=0: amt3to6=0: over7to12=0: amt7to12=0: over12=0: amt12=0: totalLoanAmount=0: totalShareAmount=0
	do while not rs.eof
		mOver=rs("overdue")
		if mOver<3 then
			if rs("thisLoanBal"&acPeriod)>0 then
				over1to2=over1to2+1
				amt1to2=amt1to2+rs("thisLoanBal"&acPeriod)
			end if
		elseif mOver<7 then
			over3to6=over3to6+1
			amt3to6=amt3to6+rs("thisLoanBal"&acPeriod)
		elseif mOver<13 then
			over7to12=over7to12+1
			amt7to12=amt7to12+rs("thisLoanBal"&acPeriod)
		else
			over12=over12+1
			amt12=amt12+rs("thisLoanBal"&acPeriod)
		end if
		totalLoanAmount=totalLoanAmount+rs("thisLoanBal"&acPeriod)
		totalShareAmount=totalShareAmount+rs("thisShrBal"&acPeriod)
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write right(spaces&formatnumber(rs("thisLoanBal"&acPeriod),2),20)
		objFile.Write right(spaces&mOver,20)
		objFile.Write right(spaces&formatnumber(rs("thisShrBal"&acPeriod),2),20)
		objFile.WriteLine ""
		rs.movenext
	loop
	for idx = 1 to 120
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"2 months or under : ",80)
	objFile.Write right(spaces&over1to2,20)
	objFile.Write right(spaces&formatnumber(amt1to2,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"3 to 6 months : ",80)
	objFile.Write right(spaces&over3to6,20)
	objFile.Write right(spaces&formatnumber(amt3to6,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"7 to 12 months : ",80)
	objFile.Write right(spaces&over7to12,20)
	objFile.Write right(spaces&formatnumber(amt7to12,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"Over 12 months : ",80)
	objFile.Write right(spaces&over12,20)
	objFile.Write right(spaces&formatnumber(amt12,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"Total loan amount : ",80)
	objFile.Write right(spaces&spaces,20)
	objFile.Write right(spaces&formatnumber(totalLoanAmount,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"Total share amount : ",80)
	objFile.Write right(spaces&spaces,20)
	objFile.Write right(spaces&formatnumber(totalShareAmount,2),20)
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
<title>Delinquent Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="5">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="5">Delinquent Report</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="200"><b>Name</b></td>
		<td width="130" align="right"><b>Loan Amount</b></td>
		<td width="130" align="right"><b>Over-due</b></td>
		<td width="130" align="right"><b>Share Amount</b></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>
<%
over1to2=0: amt1to2=0: over3to6=0: amt3to6=0: over7to12=0: amt7to12=0: over12=0: amt12=0: totalLoanAmount=0: totalShareAmount=0
do while not rs.eof
	mOver=rs("overdue")
	if mOver<3 then
		if rs("thisLoanBal"&acPeriod)>0 then
			over1to2=over1to2+1
			amt1to2=amt1to2+rs("thisLoanBal"&acPeriod)
		end if
	elseif mOver<7 then
		over3to6=over3to6+1
		amt3to6=amt3to6+rs("thisLoanBal"&acPeriod)
	elseif mOver<13 then
		over7to12=over7to12+1
		amt7to12=amt7to12+rs("thisLoanBal"&acPeriod)
	else
		over12=over12+1
		amt12=amt12+rs("thisLoanBal"&acPeriod)
	end if
	totalLoanAmount=totalLoanAmount+rs("thisLoanBal"&acPeriod)
	totalShareAmount=totalShareAmount+rs("thisShrBal"&acPeriod)
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td align="right"><%=formatNumber(rs("thisLoanBal"&acPeriod),2)%></td>
		<td align="right"><%=mOver%></td>
		<td align="right"><%=formatNumber(rs("thisShrBal"&acPeriod),2)%></td>
	</tr>
<%
	rs.movenext
loop
%>
	<tr><td colspan=5><hr></td></tr>
	<tr height="20">
		<td colspan="3" class="b10" align="right">2 months or under : </td>
		<td align="right"><%=over1to2%></td>
		<td align="right"><%=formatNumber(amt1to2,2)%></td>
	</tr>
	<tr height="20">
		<td colspan="3" class="b10" align="right">3 to 6 months : </td>
		<td align="right"><%=over3to6%></td>
		<td align="right"><%=formatNumber(amt3to6,2)%></td>
	</tr>
	<tr height="20">
		<td colspan="3" class="b10" align="right">7 to 12 months : </td>
		<td align="right"><%=over7to12%></td>
		<td align="right"><%=formatNumber(amt7to12,2)%></td>
	</tr>
	<tr height="20">
		<td colspan="3" class="b10" align="right">Over 12 months : </td>
		<td align="right"><%=over12%></td>
		<td align="right"><%=formatNumber(amt12,2)%></td>
	</tr>
	<tr height="20">
		<td colspan="3" class="b10" align="right">Total loan amount : </td>
		<td></td>
		<td align="right"><%=formatNumber(totalLoanAmount,2)%></td>
	</tr>
	<tr height="20">
		<td colspan="3" class="b10" align="right">Total share amount : </td>
		<td></td>
		<td align="right"><%=formatNumber(totalShareAmount,2)%></td>
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
