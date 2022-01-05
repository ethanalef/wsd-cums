<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
memNo = request("memNo")
mMonth1 = left(request("mPeriod1"),2)
mYear1 = right(request("mPeriod1"),4)
mMonth2 = left(request("mPeriod2"),2)
mYear2 = right(request("mPeriod2"),4)
sortBy = request("sortBy")

if mMonth1="" or mYear1="" or mMonth2="" or mYear2="" or sortBy="" then
	response.redirect "acTxList.asp"
end if

mDate1 = mYear1&"/"&mMonth1&"/1"
mDate2 = mYear2&"/"&mMonth2&"/"&day(dateAdd("m",1,mYear2&"/"&mMonth2&"/1")-1)

SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")

if memNo<>"" then
	sql_filter = " and a.memNo="&memNo
end if

sql = "select a.*,b.memName from memTx a,memMaster b where a.deleted=0 and b.deleted=0 and a.memNo=b.memNo"&sql_filter&" and a.txDate between '"&mDate1&"' and '"&mDate2&"' order by a."&sortBy

'response.write sql
'response.end

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
	objFile.Write "A/C Transaction List"
	objFile.WriteLine ""
	objFile.Write "Account Year : "&m&"  Period : "&y
	objFile.WriteLine ""
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write left("Tran. No."&spaces,10)
	objFile.Write left("Date"&spaces,11)
	objFile.Write left("Tre. No."&spaces,10)
	objFile.Write right(spaces&"Share Paid In",20)
	objFile.Write right(spaces&"Share Withdrawn",20)
	objFile.Write right(spaces&"Loan Paid",20)
	objFile.Write right(spaces&"Interest Paid",20)
	objFile.Write right(spaces&"Amount Loaded",20)
	objFile.Write right(spaces&"League Due",20)
	objFile.Write right(spaces&"Total Amount",20)
	objFile.WriteLine ""
	for idx = 1 to 231
		objFile.Write "-"
	next
	objFile.WriteLine ""
	sharePaid=0
	shareWithdrawn=0
	amtLoan=0
	loanPaid=0
	interestPaid=0
	leagueDue=0
	m99=0
	mAT=0
	mSD=0
	mAD=0
	do while not rs.eof
		totalAmount=(rs("sharePaid")+rs("interestPaid")+rs("loanPaid"))-(rs("shareWithdrawn")+rs("amtLoan"))
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write left(rs("memTxNo")&spaces,10)
		objFile.Write left(right("0"&day(rs("txDate")),2)&"/"&right("0"&month(rs("txDate")),2)&"/"&year(rs("txDate"))&spaces,11)
		objFile.Write left(rs("treNo")&spaces,10)
		objFile.Write right(spaces&formatnumber(rs("sharePaid"),2),20)
		objFile.Write right(spaces&formatnumber(rs("shareWithdrawn"),2),20)
		objFile.Write right(spaces&formatnumber(rs("loanPaid"),2),20)
		objFile.Write right(spaces&formatnumber(rs("interestPaid"),2),20)
		objFile.Write right(spaces&formatnumber(rs("amtLoan"),2),20)
		objFile.Write right(spaces&formatnumber(0,2),20)
		objFile.Write right(spaces&formatnumber(totalAmount,2),20)
		objFile.WriteLine ""
		sharePaid=sharePaid+rs("sharePaid")
		shareWithdrawn=shareWithdrawn+rs("shareWithdrawn")
		loanPaid=loanPaid+rs("loanPaid")
		interestPaid=interestPaid+rs("interestPaid")
		amtLoan=amtLoan+rs("amtLoan")
		leagueDue=leagueDue+0
		if rs("treNo")="99" then m99=m99+totalAmount end if
		if rs("treNo")="AT" then mAT=mAT+totalAmount end if
		if rs("treNo")="SD" then mSD=mSD+totalAmount end if
		if rs("treNo")="AD" then mAD=mAD+totalAmount end if
		rs.movenext
	loop
	for idx = 1 to 231
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write right(spaces&"Total Share Paid In :",50)
	objFile.Write right(spaces&formatnumber(sharePaid,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Total Share Withdrawn :",50)
	objFile.Write right(spaces&formatnumber(shareWithdrawn,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Total Amount Loaned :",50)
	objFile.Write right(spaces&formatnumber(amtLoan,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Total loan Paid :",50)
	objFile.Write right(spaces&formatnumber(loanPaid,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Total Interest Paid :",50)
	objFile.Write right(spaces&formatnumber(interestPaid,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Total League Due :",50)
	objFile.Write right(spaces&formatnumber(leagueDue,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Treasurer 99 Pay In :",50)
	objFile.Write right(spaces&formatnumber(m99,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Treasurer AT Pay In :",50)
	objFile.Write right(spaces&formatnumber(mAT,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Treasurer SD Pay In :",50)
	objFile.Write right(spaces&formatnumber(mSD,2),20)
	objFile.WriteLine ""
	objFile.Write right(spaces&"Treasurer AD Pay In :",50)
	objFile.Write right(spaces&formatnumber(mAD,2),20)
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
<title>A/C Transaction List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="12"><font size="4">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="12"><font size="4">A/C Transaction List</font></td>
	</tr>
	<tr height="30" valign="top" align="center">
		<td colspan="12">Account Year : <%=y%> &nbsp; &nbsp; Period : <%=m%></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="150"><b>Name</b></td>
		<td width="80"><b>Tran. No.</b></td>
		<td width="100"><b>Date</b></td>
		<td width="50"><b>Tre. No.</b></td>
		<td width="130" align=right><b>Share Paid In</b></td>
		<td width="130" align=right><b>Share Withdrawn</b></td>
		<td width="130" align=right><b>Loan Paid</b></td>
		<td width="130" align=right><b>Interest Paid</b></td>
		<td width="130" align=right><b>Amount Loaded</b></td>
		<td width="130" align=right><b>League Due</b></td>
		<td width="130" align=right><b>Total Amount</b></td>
	</tr>
	<tr><td colspan=12><hr></td></tr>
<%
sharePaid=0
shareWithdrawn=0
amtLoan=0
loanPaid=0
interestPaid=0
leagueDue=0
m99=0
mAT=0
mSD=0
mAD=0
do while not rs.eof
	totalAmount=(rs("sharePaid")+rs("interestPaid")+rs("loanPaid"))-(rs("shareWithdrawn")+rs("amtLoan"))
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td><%=rs("memTxNo")%></td>
		<td><%=right("0"&day(rs("txDate")),2)&"/"&right("0"&month(rs("txDate")),2)&"/"&year(rs("txDate"))%></td>
		<td><%=rs("treNo")%></td>
		<td align=right><%=formatNumber(rs("sharePaid"),2)%></td>
		<td align=right><%=formatNumber(rs("shareWithdrawn"),2)%></td>
		<td align=right><%=formatNumber(rs("loanPaid"),2)%></td>
		<td align=right><%=formatNumber(rs("interestPaid"),2)%></td>
		<td align=right><%=formatNumber(rs("amtLoan"),2)%></td>
		<td align=right><%=formatNumber(0,2)%></td>
		<td align=right><%=formatNumber(totalAmount,2)%></td>
<%
	sharePaid=sharePaid+rs("sharePaid")
	shareWithdrawn=shareWithdrawn+rs("shareWithdrawn")
	loanPaid=loanPaid+rs("loanPaid")
	interestPaid=interestPaid+rs("interestPaid")
	amtLoan=amtLoan+rs("amtLoan")
	leagueDue=leagueDue+0
	if rs("treNo")="99" then m99=m99+totalAmount end if
	if rs("treNo")="AT" then mAT=mAT+totalAmount end if
	if rs("treNo")="SD" then mSD=mSD+totalAmount end if
	if rs("treNo")="AD" then mAD=mAD+totalAmount end if
	rs.movenext
loop
%>
	<tr><td colspan=12><hr></td></tr>
	<tr>
		<td colspan="4" class="b10" align="right">Total Share Paid In : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(sharePaid,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Total Share Withdrawn : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(shareWithdrawn,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Total Amount Loaned : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(amtLoan,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Total loan Paid : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(loanPaid,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Total Interest Paid : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(interestPaid,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Total League Due : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(leagueDue,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Treasurer 99 Pay In : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(m99,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Treasurer AT Pay In : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(mAT,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Treasurer SD Pay In : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(mSD,2)%></td>
	</tr>
	<tr>
		<td colspan="4" class="b10" align="right">Treasurer AD Pay In : </td>
		<td></td>
		<td  colspan="7"><%=formatNumber(mAD,2)%></td>
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
