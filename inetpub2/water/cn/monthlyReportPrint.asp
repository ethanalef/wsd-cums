<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
id=request("id")

sql = "select * from monthlyReport where uid="&id
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3, 3
startDate = year(rs("StartDate"))&"/"&month(rs("StartDate"))&"/"&day(rs("StartDate"))
endDate = year(rs("EndDate"))&"/"&month(rs("endDate"))&"/"&day(rs("endDate"))

set bdrs = conn.execute("select sum(interview) from meetingNotes where rpDate between '"&startDate&"' and '"&endDate&"'")
if isnull(bdrs(0)) then interview=0 else interview=bdrs(0) end if
bdrs.close

set bdrs = conn.execute("select count(*) from meetingNotes where rpDate between '"&startDate&"' and '"&endDate&"'")
if isnull(bdrs(0)) then meetingCount=0 else meetingCount=bdrs(0) end if
bdrs.close

sql = "select count(*),sum(loanAmt) from meetingNotes0 a, loanApp b, meetingNotes c where a.appId=b.uid and a.rpId=c.uid and c.rpDate between '"&startDate&"' and '"&endDate&"'"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then LoanNo=0 else LoanNo=bdrs(0) end if
if isnull(bdrs(1)) then LoanAmt=0 else LoanAmt=bdrs(1) end if
bdrs.close

sql = "select count(*),sum(loanAmt) from meetingNotes1 a, loanApp b, meetingNotes c where a.appId=b.uid and a.rpId=c.uid and c.rpDate between '"&startDate&"' and '"&endDate&"'"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then rejectNo=0 else rejectNo=bdrs(0) end if
if isnull(bdrs(1)) then rejectAmt=0 else rejectAmt=bdrs(1) end if
bdrs.close

sql = "select count(*),sum(loanAmt) from meetingNotes0 a, loanApp b, meetingNotes c where a.appId=b.uid and a.rpId=c.uid and b.uid in (select c.loanAppID from loanReason c, reason d where c.reasonID=d.uid and d.reasonType=1) and c.rpDate between '"&startDate&"' and '"&endDate&"'"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then totalNo1=0 else totalNo1=bdrs(0) end if
if isnull(bdrs(1)) then totalLoanAmt1=0 else totalLoanAmt1=bdrs(1) end if
bdrs.close

sql = "select count(*),sum(loanAmt) from meetingNotes0 a, loanApp b, meetingNotes c where a.appId=b.uid and a.rpId=c.uid and b.uid in (select c.loanAppID from loanReason c, reason d where c.reasonID=d.uid and d.reasonType<>1) and c.rpDate between '"&startDate&"' and '"&endDate&"'"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then totalNo2=0 else totalNo2=bdrs(0) end if
if isnull(bdrs(1)) then totalLoanAmt2=0 else totalLoanAmt2=bdrs(1) end if
bdrs.close

%>
<html>
<head>
<title>Monthly Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="top" align="center">
		<td width="700">
			<font size="5">���q�u�{�p�x�W���U���U�کe���|�V���Ʒ|���i��</font><br>
			<br>
			<font size="4">����G <%=year(rs("startDate"))%> �~ <%=month(rs("startDate"))%> �� <%=day(rs("startDate"))%> �� ��
			<%=year(rs("endDate"))%> �~ <%=month(rs("endDate"))%> �� <%=day(rs("endDate"))%> ��</font>
		</td>
	</tr>
	<tr valign="top">
		<td width="700">
			<font size="4">
			<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �e �� �| �b �W �C �� �� �A �� �b �E �s �� �� �B �l �} �| ĳ <%=meetingCount%> <br>
			�� �C �e �� <% if rs("absent")=0 then response.write "�L" else response.write "��" end if %> �s �� �T �� �� �X �u �A �� �e �� �| �w �� �� �U �C �� �� �G<br>
			<%=rs("actions")%><br>
			<br>
			�� �� �U �� �� �� �@ <%=LoanNo+rejectNo%> �W �C<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			�� �� <%=LoanNo%> �W �� �� �U �� �C �@ $ <%=formatnumber(LoanAmt,2)%><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			�� �� <%=rejectNo%> �W �� �� �U �� �C �@ $ <%=formatnumber(rejectAmt,2)%><br>
			<br>
			�� �� �U �� �� �z �� �� �] �� �z �^ �G <%=rs("rejectReason")%><br>
			<br>
			�U �� �� �~ �G <br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			�� �I �� �� �� �� �� <%=totalNo1%>  �W �C �@ �U �X $ <%=formatnumber(totalLoanAmt1,2)%><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			��&nbsp; &nbsp;��&nbsp; &nbsp;��&nbsp; &nbsp;�~&nbsp; &nbsp;�� <%=totalNo2%>  �W �C �@ �U �X  $ <%=formatnumber(totalLoanAmt2,2)%><br>
			<br>
			�� �� �� �� �� �� �� �� �@ �C <br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			�� �a  �W �� �� �� �� �U �� �C<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			�� ��  �W �� �� �� �� �U �� �C<br>
			<br>
			�� �e �� �| �� �� �� �W �C �� �U �� �� �� <%=interview%> �W �C<br>
			<br>
			�� �L �� �� �G<%=rs("others")%><br>
			<br><br><br>
			ñ �p �G ________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ________________________ <br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �D �u (���è})
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �� �� (��ģ��)
			<br><br>
			</font>
		</td>
	</tr>
	<tr valign="top">
		<td width="700" align="center">
			<font size="4">
			�� ���G <%=year(rs("rpDate"))%> �~ <%=month(rs("rpDate"))%> �� <%=day(rs("rpDate"))%> ��</font>
			</font>
		</td>
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