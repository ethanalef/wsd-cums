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
			<font size="5">機電工程署儲蓄互助社貸款委員會向董事會報告書</font><br>
			<br>
			<font size="4">日期： <%=year(rs("startDate"))%> 年 <%=month(rs("startDate"))%> 月 <%=day(rs("startDate"))%> 日 至
			<%=year(rs("endDate"))%> 年 <%=month(rs("endDate"))%> 月 <%=day(rs("endDate"))%> 日</font>
		</td>
	</tr>
	<tr valign="top">
		<td width="700">
			<font size="4">
			<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本 委 員 會 在 上 列 日 期 ， 曾 在 九 龍 辦 事 處 召 開 會 議 <%=meetingCount%> <br>
			次 。 委 員 <% if rs("absent")=0 then response.write "無" else response.write "有" end if %> 連 續 三 次 不 出 席 ， 本 委 員 會 已 採 取 下 列 行 動 ：<br>
			<%=rs("actions")%><br>
			<br>
			申 請 貸 款 社 員 共 <%=LoanNo+rejectNo%> 名 。<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			批 準 <%=LoanNo%> 名 社 員 貸 款 。 共 $ <%=formatnumber(LoanAmt,2)%><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			拒 絕 <%=rejectNo%> 名 社 員 貸 款 。 共 $ <%=formatnumber(rejectAmt,2)%><br>
			<br>
			拒 絕 貸 款 之 理 由 為 （ 略 述 ） ： <%=rs("rejectReason")%><br>
			<br>
			貸 款 用 途 ： <br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			應 付 不 時 之 需 者 <%=totalNo1%>  名 。 共 貸 出 $ <%=formatnumber(totalLoanAmt1,2)%><br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			生&nbsp; &nbsp;產&nbsp; &nbsp;用&nbsp; &nbsp;途&nbsp; &nbsp;者 <%=totalNo2%>  名 。 共 貸 出  $ <%=formatnumber(totalLoanAmt2,2)%><br>
			<br>
			申 請 延 期 還 款 社 員 共 。 <br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			批 淮  名 社 員 延 期 貸 款 。<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			拒 絕  名 社 員 延 期 貸 款 。<br>
			<br>
			本 委 員 會 曾 接 見 上 列 之 貸 款 社 員 <%=interview%> 名 。<br>
			<br>
			其 他 事 項 ：<%=rs("others")%><br>
			<br><br><br>
			簽 署 ： ________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ________________________ <br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 主 席 (任永良)
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 秘 書 (李耀威)
			<br><br>
			</font>
		</td>
	</tr>
	<tr valign="top">
		<td width="700" align="center">
			<font size="4">
			日 期： <%=year(rs("rpDate"))%> 年 <%=month(rs("rpDate"))%> 月 <%=day(rs("rpDate"))%> 日</font>
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