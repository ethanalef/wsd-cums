<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="navigator.asp" -->
<%
sql = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close
if acPeriod<=4 then
	m=acPeriod+8
	y=acYear
else
	m=acPeriod-4
	y=acYear+1
end if

if request("del")<>"" then
	memTxNo = request("del")

	if month(date())=m and year(date())=y then
		mDate = y&"/"&m&"/"&mDay
	else
		if m=12 then
			mDate = cdate(y&"/12/31")
		else
			mDate = cdate(y&"/"&m+1&"/1")-1
		end if
		mDate = year(mDate)&"/"&month(mDate)&"/"&day(mDate)
	end if

	conn.begintrans
	sql = "select * from memTx where memTxNo="&memTxNo
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn, 2, 2
	For Each Field in rs.fields
		TheString = Field.name & "= rs(""" & Field.name & """)"
		Execute(TheString)
	Next
	rs("deleted") =-1
	rs.update
	rs.close

	if sharePaid<>"" then
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&sharePaid&" where glId='0205'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Delete Share Paid','"&mDate&"','C',"&sharePaid&",0 from glTx")
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&sharePaid&" where glId='0401'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0401','Delete Share Paid','"&mDate&"','D',"&sharePaid&",0 from glTx")
		conn.execute("update memMaster set thisShrBal"&acPeriod&"=thisShrBal"&acPeriod&" - "&sharePaid)
	end if
	if shareWithdrawn<>"" then
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&shareWithdrawn&" where glId='0205'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Delete Share Withdrawn','"&mDate&"','D',"&shareWithdrawn&",0 from glTx")
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&shareWithdrawn&" where glId='0401'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0401','Delete Share Withdrawn','"&mDate&"','C',"&shareWithdrawn&",0 from glTx")
		conn.execute("update memMaster set thisShrBal"&acPeriod&"=thisShrBal"&acPeriod&" + "&shareWithdrawn)
	end if
	if amtLoan<>"" then
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&amtLoan&" where glId='0205'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Delete Amount Loaned','"&mDate&"','D',"&amtLoan&",0 from glTx")
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&amtLoan&" where glId='0201'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delete Amount Loaned','"&mDate&"','C',"&amtLoan&",0 from glTx")
		conn.execute("update memMaster set thisLoanBal"&acPeriod&"=thisLoanBal"&acPeriod&" - "&amtLoan)
	end if
	if interestPaid<>"" then
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&interestPaid&" where glId='0205'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Delete Interest Paid','"&mDate&"','C',"&interestPaid&",0 from glTx")
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&interestPaid&" where glId='0501'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0501','Delete Interest Paid','"&mDate&"','D',"&interestPaid&",0 from glTx")
	end if
	if loanPaid<>"" then
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&loanPaid&" where glId='0205'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Delete Load Paid','"&mDate&"','C',"&loanPaid&",0 from glTx")
		conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&loanPaid&" where glId='0201'")
		conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delete Load Paid','"&mDate&"','D',"&loanPaid&",0 from glTx")
		conn.execute("update memMaster set thisLoanBal"&acPeriod&"=thisLoanBal"&acPeriod&" + "&loanPaid)
	end if

	addUserLog "Delete Account transaction"
	msg = glId&" deleted"

	conn.committrans
end if

if request("year") = "" then
	mYear = y
else
	mYear = request("year")
end if
if request("month") = "" then
	mMonth = m
else
	mMonth = request("month")
end if

set rs = server.createobject("ADODB.Recordset")
sql = "select * from memTx where deleted=0 and month(txDate)="&mMonth&" and year(txDate)="&mYear&" order by memTxNo desc"
rs.open sql, conn, 3

if rs.eof then
	response.redirect "acTxDetail.asp"
end if
%>
<html>
<head>
<title>個人賬入數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<center>
<form method="post" action="acTx.asp" name="form1">
Month:
<select name="month">
<%
for idx = 12 to 1 step -1
	if idx = cint(mMonth) then
		response.write "<option selected>"&idx&"</option>"
	else
		response.write "<option>"&idx&"</option>"
	end if
next
%>
</select>
Year:
<select name="year">
<%
set rs1 = conn.execute("select min(year(txDate)) from memTx where deleted=0")
for idx = y to rs1(0) step -1
	if idx = cint(mYear) then
		response.write "<option selected>"&idx&"</option>"
	else
		response.write "<option>"&idx&"</option>"
	end if
next
%>
</select>
<input type="submit" value="Submit" name="submit" class="sbttn">
</form>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font color="#FFFFFF">序號</font></td>
	<td><font color="#FFFFFF">日期</font></td>
	<td><font color="#FFFFFF">社員號碼</font></td>
	<td><font color="#FFFFFF">類別</font></td>
	<td><font color="#FFFFFF">存款</font></td>
	<td><font color="#FFFFFF">退股</font></td>
	<td><font color="#FFFFFF">貸款</font></td>
	<td><font color="#FFFFFF">每月還款</font></td>
	<td><font color="#FFFFFF">貸款利息</font></td>
	<td><font color="#FFFFFF">還款</font></td>
<%if session("userLevel")<>5 then%>
	<td bgcolor="#FFFFFF"><a href="acTxDetail.asp">新增</a></td>
<%end if%>
  </tr>
<%
do while not rs.eof
%>
  <tr bgcolor="#FFFFFF">
	<td align="right"><%=rs("memTxNo")%></font></td>
	<td><%=right("0"&day(rs("txDate")),2)&"/"&right("0"&month(rs("txDate")),2)&"/"&year(rs("txDate"))%></td>
	<td align="right"><%=rs("memNo")%></font></td>
	<td align="center"><%=rs("treNo")%></font></td>
	<td align="right"><%if rs("sharePaid")<>0 then response.write formatNumber(rs("sharePaid"),2) end if%></td>
	<td align="right"><%if rs("shareWithdrawn")<>0 then response.write formatNumber(rs("shareWithdrawn"),2) end if%></td>
	<td align="right"><%if rs("amtLoan")<>0 then response.write formatNumber(rs("amtLoan"),2) end if%></td>
	<td align="right"><%if rs("monthlyRepaid")<>0 then response.write formatNumber(rs("monthlyRepaid"),2) end if%></td>
	<td align="right"><%if rs("interestPaid")<>0 then response.write formatNumber(rs("interestPaid"),2) end if%></td>
	<td align="right"><%if rs("loanPaid")<>0 then response.write formatNumber(rs("loanPaid"),2) end if%></td>
<%if session("userLevel")<>5 then%>
	<td><a href="acTx.asp?del=<%=rs("memTxNo")%>" onclick="return confirm('Delete this record?')">刪除</a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
