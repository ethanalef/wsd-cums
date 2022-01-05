<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQl = "select * from memMaster where deleted=0 and salaryDedut<>0 order by memNo"
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
	objFile.Write "WSDS Credit Union"
	objFile.WriteLine ""
	objFile.Write "A/C Check List for Salary Deduction"
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write right(spaces&"Total Amount",20)
	objFile.Write right(spaces&"Loan Paid",20)
	objFile.Write right(spaces&"Interest",20)
	objFile.Write right(spaces&"Share",10)
	objFile.WriteLine ""
	for idx = 1 to 130
		objFile.Write "-"
	next
	objFile.WriteLine ""
	loanRepaid=0: OSInterest=0: totalAmount=0: totalShare=0
	do while not rs.eof
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write right(spaces&formatnumber(int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10,2),20)
		objFile.Write right(spaces&formatnumber(rs("loanRepaid"),2),20)
		objFile.Write right(spaces&formatnumber(rs("OSInterest"),2),20)
		objFile.Write right(spaces&formatNumber((int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10)-rs("loanRepaid")-rs("OSInterest"),2),10)
		objFile.WriteLine ""
		loanRepaid=loanRepaid+rs("loanRepaid")
		OSInterest=OSInterest+rs("OSInterest")
		totalAmount=totalAmount+(int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10)
		totalShare=totalShare+((int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10)-rs("loanRepaid")-rs("OSInterest"))
		rs.movenext
	loop
	for idx = 1 to 130
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"Total : ",60)
	objFile.Write right(spaces&formatnumber(totalAmount,2),20)
	objFile.Write right(spaces&formatnumber(loanRepaid,2),20)
	objFile.Write right(spaces&formatnumber(OSInterest,2),20)
	objFile.Write right(spaces&formatnumber(totalShare,2),10)
	objFile.WriteLine ""
	objFile.Write right(spaces&spaces&"Total Staff : ",50)
	objFile.Write rs.recordcount
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
<title>A/C Check List for Salary Deduction</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="5">EWSDS Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="5">A/C Check List for Salary Deduction</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="200"><b>Name</b></td>
		<td width="130" align="right"><b>Total Amount</b></td>
		<td width="130" align="right"><b>Loan Paid</b></td>
		<td width="130" align="right"><b>Interest</b></td>
		<td width="80" align="right"><b>Share</b></td>
	</tr>
	<tr><td colspan=6><hr></td></tr>
<%
loanRepaid=0: OSInterest=0: totalAmount=0: totalShare=0
do while not rs.eof
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td align="right"><%=formatNumber(int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10,2)%></td>
		<td align="right"><%=formatNumber(rs("loanRepaid"),2)%></td>
		<td align="right"><%=formatNumber(rs("OSInterest"),2)%></td>
		<td align="right"><%=formatNumber((int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10)-rs("loanRepaid")-rs("OSInterest"),2)%></td>
	</tr>
<%
	loanRepaid=loanRepaid+rs("loanRepaid")
	OSInterest=OSInterest+rs("OSInterest")
	totalAmount=totalAmount+(int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10)
	totalShare=totalShare+((int((rs("loanRepaid")+rs("OSInterest")+9.99)/10)*10)-rs("loanRepaid")-rs("OSInterest"))
	rs.movenext
loop
%>
	<tr><td colspan=6><hr></td></tr>
	<tr>
		<td colspan="2" class="b10" align="right">Total : </td>
		<td align="right"><%=formatNumber(totalAmount,2)%></td>
		<td align="right"><%=formatNumber(loanRepaid,2)%></td>
		<td align="right"><%=formatNumber(OSInterest,2)%></td>
		<td align="right"><%=formatNumber(totalShare,2)%></td>
	</tr>
	<tr>
		<td colspan="2" class="b10" align="right">Total Staff : </td>
		<td align="right"><%=rs.recordcount%></td>
		<td colspan="2"></td>
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
