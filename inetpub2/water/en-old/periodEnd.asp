<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request("process")<>"" then
	conn.begintrans
	sql = "select * from glControl"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn, 2, 2
	acPeriod=rs("acPeriod")
	acYear=rs("acYear")
	if acPeriod<=4 then
		m=acPeriod+8
		y=acYear
	else
		m=acPeriod-4
		y=acYear+1
	end if
	if m=12 then
		dayOfMonth = 31
	else
		dayOfMonth = day(cdate(y&"/"&m+1&"/1")-1)
	end if
	conn.execute("update memMaster set OSInterest=0,loanRepaid=0,salaryDedut=0 where thisLoanBal"&rs("acPeriod")&"=0")
	sql = "select a.thisInterest,a.OSinterest,a.loanRepaid,a.salaryDedut,b.txDate,b.amtLoan,b.monthlyRepaid,b.calcInterest from memMaster a, memTx b where a.memNo=b.memNo and a.deleted=0 and month(b.txDate)="&m&" and year(b.txDate)="&y&" and b.amtLoan>0"
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.open sql, conn, 2, 2
	do while not rs2.eof
		rs2("loanRepaid")=rs2("monthlyRepaid")
		if rs2("calcInterest")=0 then
			rs2("OSinterest")=0
			rs2("salaryDedut")=rs2("loanRepaid")
			rs2("thisInterest")=0
		else
			OSinterest=int((((rs2("amtLoan")+rs2("amtLoan")/(rs2("amtLoan")/rs2("monthlyRepaid")))/200+(rs2("amtLoan")/100/dayOfMonth*(dayOfMonth-day(rs2("txDate")))/(rs2("amtLoan")/rs2("monthlyRepaid"))))+0.04)*20)/20
			rs2("OSinterest")=OSinterest
			rs2("salaryDedut")=int((rs2("OSinterest")+rs2("loanRepaid")+9.99)/10)*10
			rs2("thisInterest")=int(((dayOfMonth-day(rs2("txDate")))/dayOfMonth*rs2("amtLoan")*0.01)+0.99)
		end if
		rs2.update
		rs2.movenext
	loop
	rs2.close
	set rs2=nothing

'' ***** update thisinterest
	conn.execute("update memMaster set thisInterest=(thisLoanBal"&acPeriod&"*0.01)-OSinterest+thisInterest where deleted=0 and thisLoanBal"&acPeriod&">0 and calcInterest=1")
	conn.execute("update memMaster set thisInterest=0 where (deleted=0 and thisLoanBal"&acPeriod&"=0) or calcInterest=0")

	if rs("acPeriod")=1 then
		conn.execute("update memMaster set overdue=overdue+1 where lastLoanBal12=thisLoanBal1 and thisloanBal1<>0 and deleted=0")
	else
		conn.execute("update memMaster set overdue=overdue+1 where thisLoanBal"&rs("acPeriod")-1&"=thisLoanBal"&rs("acPeriod")&" and thisLoanBal"&rs("acPeriod")&"<>0 and deleted=0")
	end if

	conn.execute("update memMaster set autopayAmt=autopayPerm where deleted=0")

	if rs("acPeriod")=12 then
		rs("acPeriod") = 1
		rs("acYear") = rs("acYear") + 1
		conn.execute("update glMaster set openBal=lastBal12 where deleted=0")
		conn.execute("update glMaster set lastBal1=thisBal1, lastBal2=thisBal2, lastBal3=thisBal3, lastBal4=thisBal4, lastBal5=thisBal5, lastBal6=thisBal6, lastBal7=thisBal7, lastBal8=thisBal8, lastBal9=thisBal9, lastBal10=thisBal10, lastBal11=thisBal11, lastBal12=thisBal12 where deleted=0")
		conn.execute("update glMaster set thisBal1=thisBal12, thisBal2=0, thisBal3=0, thisBal4=0, thisBal5=0, thisBal6=0, thisBal7=0, thisBal8=0, thisBal9=0, thisBal10=0, thisBal11=0, thisBal12=0  where deleted=0")
		conn.execute("update memMaster set lastShrBal1=thisShrBal1, lastShrBal2=thisShrBal2, lastShrBal3=thisShrBal3, lastShrBal4=thisShrBal4, lastShrBal5=thisShrBal5, lastShrBal6=thisShrBal6, lastShrBal7=thisShrBal7, lastShrBal8=thisShrBal8, lastShrBal9=thisShrBal9, lastShrBal10=thisShrBal10, lastShrBal11=thisShrBal11, lastShrBal12=thisShrBal12 where deleted=0")
		conn.execute("update memMaster set lastLoanBal1=thisLoanBal1, lastLoanBal2=thisLoanBal2, lastLoanBal3=thisLoanBal3, lastLoanBal4=thisLoanBal4, lastLoanBal5=thisLoanBal5, lastLoanBal6=thisLoanBal6, lastLoanBal7=thisLoanBal7, lastLoanBal8=thisLoanBal8, lastLoanBal9=thisLoanBal9, lastLoanBal10=thisLoanBal10, lastLoanBal11=thisLoanBal11, lastLoanBal12=thisLoanBal12 where deleted=0")
		conn.execute("update memMaster set thisShrBal1=thisShrBal12, thisShrBal2=0, thisShrBal3=0, thisShrBal4=0, thisShrBal5=0, thisShrBal6=0, thisShrBal7=0, thisShrBal8=0, thisShrBal9=0, thisShrBal10=0, thisShrBal11=0, thisShrBal12=0  where deleted=0")
		conn.execute("update memMaster set thisLoanBal1=thisLoanBal12, thisLoanBal2=0, thisLoanBal3=0, thisLoanBal4=0, thisLoanBal5=0, thisLoanBal6=0, thisLoanBal7=0, thisLoanBal8=0, thisLoanBal9=0, thisLoanBal10=0, thisLoanBal11=0, thisLoanBal12=0  where deleted=0")
	else
		conn.execute("update memMaster set thisShrBal"&rs("acPeriod")+1&"=thisShrBal"&rs("acPeriod")&", thisLoanBal"&rs("acPeriod")+1&"=thisLoanBal"&rs("acPeriod"))
		conn.execute("update glMaster set thisBal"&rs("acPeriod")+1&"=thisBal"&rs("acPeriod"))
		rs("acPeriod") = rs("acPeriod") + 1
	end if
	rs.update

	addUserLog "Period End process"
	conn.committrans

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>每月完結</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>每月完結</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>">
<input type="submit" name="process" value="確定">
</form>
</center>
</body>
</html>
