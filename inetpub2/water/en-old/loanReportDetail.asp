<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "loanReport.asp"
end if

id = request("id")
set rs = server.createobject("ADODB.Recordset")
sql = "select * from glControl"
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close

sql = "select * from memMaster where memNo="&id
rs.open sql, conn
if rs.eof then
	response.redirect "loanReport.asp"
end if
For Each Field in rs.fields
	TheString = Field.name & "= rs(""" & Field.name & """)"
	Execute(TheString)
Next
rs.close
%>
<html>
<head>
<title>貸款申請列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.id.select();form1.id.focus();">
<!-- #include file="menu.asp" -->
<center>
<form name="form1" method="post" action="loanReportDetail.asp">
<table border="0" cellspacing="0" cellpadding="0">
	<tr height="30">
		<td class="b8" align="right" width="120">社員編號</td>
		<td width=10></td>
		<td width="100">
			 <input type="text" name="id" value="<%=id%>" size="4">
			<input type="submit" value="Search" name="Search" class="sbttn">
		</td>
		<td class="b8" align="right" width="50">姓名</td>
		<td width=10></td>
		<td width="200"><%=memName%></td>
		<td class="b8" align="right" width="80">職級</td>
		<td width=10></td>
		<td width="100"><%=memGrade%></td>
		<td class="b8" align="right" width="140">招聘條款</td>
		<td width=10></td>
		<td width="100"><%=employCond%></td>
	</tr>
<%
	sql = "select * from memMaster where memGuarantorNo="&memNo&" and thisLoanBal"&acPeriod&">0"
	rs.open sql, conn
	do while not rs.eof
		lastDate=DateAdd("m", rs("thisLoanBal"&acPeriod)/rs("loanRepaid") , date())
%>
	<tr height="30">
		<td class="b8" align="right">擔保他人</td>
		<td></td>
		<td><%=rs("memNo")%></td>
		<td class="b8" align="right">姓名</td>
		<td></td>
		<td><%=rs("memName")%></td>
		<td class="b8" align="right">承擔款額</td>
		<td></td>
		<td><%=rs("thisLoanBal"&acPeriod)%></td>
		<td class="b8" align="right">到期年月份</td>
		<td></td>
		<td><%=month(lastDate)&" - "&year(lastDate)%></td>
	</tr>
<%
		rs.movenext
	loop
	rs.close
%>
</table>
<br>
<table border="0" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td>
			<b>每月預定存款 :</b> <%=formatnumber(autopayPerm,2)%><br><br>
			<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
			  <tr bgcolor="#330000" align="center">
				<td><font size="2" color="#FFFFFF">月份</font></td>
				<td><font size="2" color="#FFFFFF">儲蓄結存</font></td>
			  </tr>
			<%
			m=acPeriod
			y="last"
			for idx = 11 to 0 step -1
				if this=12 then m=1:y="this" else m=m+1 end if
			%>
			  <tr bgcolor="#FFFFFF">
				<td><font size="2"><%if idx=0 then response.write "本" else response.write "前 "&idx end if%> 月</font></td>
				<td align=right><font size="2"><%=formatnumber(eval(y&"ShrBal"&m),2)%></font></td>
			  </tr>
			<%
			next
			%>
			</table>
		</td>
		<td width=10></td>
		<td>
<%
if eval("thisLoanBal"&acPeriod)>0 then
	sql = "select top 1 * from memTx where memNo="&memNo&" and amtLoan>0 order by txDate desc"
	rs.open sql, conn, 3
	if not rs.eof then
%>
			<table cellspacing="0" cellpadding="0"
			  <tr>
				<td><b>攤還期數 : </b></td><td width="70" align="center"><%=rs("amtLoan")/rs("monthlyRepaid")%></td>
				<td><b>總本金 : </b></td><td width="70" align="center"><%=rs("amtLoan")%></td>
				<td><b>總利息 : </b></td><td width="70" align="center"><%=formatnumber((rs("amtLoan")/rs("monthlyRepaid"))*OSInterest,2)%></td>
			  </tr>
			  <tr>
				<td><b>每月還款本金 : </b></td><td align="center"><%=rs("monthlyRepaid")%></td>
				<td><b>每月還款利息 : </b></td><td align="center"><%=formatnumber(OSInterest,2)%></td>
				<td colspan="2"></td>
			  </tr>
			</table>
<%
	else
		response.write "<br><br>"
	end if
else
	response.write "<br><br>"
end if
%>
			<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
			  <tr bgcolor="#330000" align="center">
				<td><font size="2" color="#FFFFFF">貸款結久本金</font></td>
				<td><font size="2" color="#FFFFFF">利息</font></td>
			  </tr>
			<%
			m=acPeriod
			y="last"
			for idx = 11 to 0 step -1
				if this=12 then m=1:y="this" else m=m+1 end if
			%>
			  <tr bgcolor="#FFFFFF">
				<td align=right><font size="2"><%=formatnumber(eval(y&"LoanBal"&m),2)%></font></td>
				<td align=right><font size="2"><%=formatnumber(eval("OSInterest"),2)%></font></td>
			  </tr>
			<%
			next
			%>
			</table>
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
<%
conn.close
set conn=nothing
%>