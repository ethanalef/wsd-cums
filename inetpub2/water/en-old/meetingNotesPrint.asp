<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
id=request("id")

SQl = "select * from meetingNotes where uid="&id
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
if isnull(rs("interview")) then interview=0 else interview=rs("interview") end if

SQl = "select count(*),sum(loanAmt) from meetingNotes1 a, loanApp b where a.appId=b.uid and a.rpId="&id
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then rejectNo=0 else rejectNo=bdrs(0) end if
if isnull(bdrs(1)) then rejectAmt=0 else rejectAmt=bdrs(1) end if
bdrs.close
SQl = "select sum(loanAmt) from meetingNotes0 a, loanApp b where a.appId=b.uid and a.rpId="&id
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then totalLoanAmt=0 else totalLoanAmt=bdrs(0) end if
bdrs.close

SQl = "select count(*),sum(loanAmt) from meetingNotes0 a, loanApp b where a.appId=b.uid and b.uid in (select c.loanAppID from loanReason c, reason d where c.reasonID=d.uid and d.reasonType=1) and a.rpId="&id
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then totalNo1=0 else totalNo1=bdrs(0) end if
if isnull(bdrs(1)) then totalLoanAmt1=0 else totalLoanAmt1=bdrs(1) end if
bdrs.close
SQl = "select count(*),sum(loanAmt) from meetingNotes0 a, loanApp b where a.appId=b.uid and b.uid in (select c.loanAppID from loanReason c, reason d where c.reasonID=d.uid and d.reasonType<>1) and a.rpId="&id
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if isnull(bdrs(0)) then totalNo2=0 else totalNo2=bdrs(0) end if
if isnull(bdrs(1)) then totalLoanAmt2=0 else totalLoanAmt2=bdrs(1) end if
bdrs.close

sql = "select b.* from meetingNotes0 a, loanApp b where a.appId=b.uid and a.rpId="&id&" order by b.memNo"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3, 3

rowPerPage=20
pageno=1
%>
<html>
<head>
<title>Meeting Notes</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<%
do while not bdrs.eof or pageno=1
%>
<table border="0" cellpadding="0" cellspacing="0" height="1000">
	<tr valign="top">
		<td colspan="2" align="center" height="40"><font size="4">機電工程署儲蓄互助社貸款委員會會議記錄</font></td>
	</tr>
	<tr valign="top">
		<td colspan="2" height="90">
			開會日期： <%=year(rs("rpDate"))%> 年 <%=month(rs("rpDate"))%> 月 <%=day(rs("rpDate"))%> 日。今年度第 <%=rs("rpNo")%> 次會議。
			第 <%=pageno%> 頁<br>
			貸款委員會 <%=rs("rpType")%> 在上列日期由 <%=rs("rpTime")%> 開始在本社辦事處舉行。<br>
			前次 <%=year(rs("lastRpDate"))%> 年 <%=month(rs("lastRpDate"))%> 月 <%=day(rs("lastRpDate"))%>
			日會議之會議記錄經讀出後通過，並加以下列修正：<br>
			<%=rs("amenment")%>
		</td>
	</tr>
	<tr valign="top" height="20">
		<td>出席委員 ：<%=rs("present")%></td><td align="right" width="100">缺席者：<%=rs("absent")%></td>
	</tr>
	<tr valign="top" height="60">
		<td colspan="2">
			出席之其他人仕：<%=rs("attendee")%><br>
			本次會議溫習章程：<%=rs("overview")%><br>
			批準下列貨款申請：<br>
		</td>
	</tr>
	<tr valign="top" height="520">
		<td colspan="2">
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>&nbsp;</td>
					<td>社員編號</td>
					<td>姓名</td>
					<td>貸款金額</td>
					<td>期數</td>
					<td>貸款理由摘要及抵押</td>
					<td>擔保人</td>
					<td>備註</td>
				</tr>
<%
thisLoanAmt=0
thisMem=0
for idx = 1 to 20
	if not bdrs.eof then%>
				<tr>
					<td align="center"><%=idx%></td>
					<td><%=bdrs("memNo")%></td>
					<td><%=bdrs("memName")%></td>
					<td align="right"><%=formatnumber(bdrs("loanAmt"),2)%></td>
					<td align="right"><%=bdrs("installment")%></td>
<%
set reasonRs = conn.execute("select reasonName from reason a,loanReason b where a.uid=b.reasonID and b.loanAppID="&bdrs("uid"))
if reasonRs.eof then
	thisReason = bdrs("otherReason1")&bdrs("otherReason2")
else
	thisReason = reasonRs.getString(,,,",")&bdrs("otherReason1")&bdrs("otherReason2")
	if right(thisReason,1)="," then thisReason=left(thisReason,len(thisReason)-1) end if
end if
%>
					<td><%if thisReason="" then response.write "&nbsp;" else response.write thisReason end if%></td>
					<td><%if isnull(bdrs("guarantorID")) or bdrs("guarantorID")=0  then response.write "&nbsp;" else response.write bdrs("guarantorID") end if%></td>
					<td><%if isnull(bdrs("remarks")) then response.write "&nbsp;" else response.write bdrs("remarks") end if%></td>
				</tr>
<%
		thisLoanAmt=thisLoanAmt+bdrs("loanAmt")
		thisMem=thisMem+1
		bdrs.movenext
	else
		response.write "<tr><td align=center>"&idx&"</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
	end if
next
%>
			</table>
		</td>
	</tr>
	<tr valign="top" height="110">
		<td colspan="2">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>此頁貸出金額</td><td align="center" width="160"><%=formatnumber(thisLoanAmt,2)%></td>
					<td>元，共</td><td align="center" width="160"><%=thisMem%></td><td>人</td>
				</tr>
				<tr>
					<td>本日記錄共</td><td align="center"><%=-int(-bdrs.recordcount/rowPerPage)%></td>
					<td>頁，總金額</td><td align="center"><%=formatnumber(totalLoanAmt,2)%></td><td>元</td>
				</tr>
				<tr>
					<td>總人數</td><td align="center"><%=bdrs.recordcount%></td>
					<td>人，共接見</td><td align="center"><%=rs("interview")%></td><td>人</td>
				</tr>
				<tr>
					<td>不時之需</td><td align="center"><%=totalNo1%></td>
					<td>人，貸出金額</td><td align="center"><%=formatnumber(totalLoanAmt1,2)%></td><td>元</td>
				</tr>
				<tr>
					<td>生產用途</td><td align="center"><%=totalNo2%></td>
					<td>人，貸出金額</td><td align="center"><%=formatnumber(totalLoanAmt2,2)%></td><td>元</td>
				</tr>
				<tr>
					<td>否決貸款申請</td><td align="center"><%=rejectNo%></td>
					<td>宗，共</td><td align="center"><%=formatnumber(rejectAmt,2)%></td><td>元</td>
				</tr>
			</table>
		<td>
	</tr>
	<tr>
		<td colspan="2">
		下次會議日期： <%=year(rs("nextRpDate"))%> 年 <%=month(rs("nextRpDate"))%> 月 <%=day(rs("nextRpDate"))%> 日，<%=rs("nextRpTime")%>于本社辦事處舉行。
		<br><br><br><br><br>
		簽署： ________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ________________________ <br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 主席 (任永良)
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 秘書 (李耀威)
		</td>
	</tr>
</table>
<%
	pageno=pageno+1
loop


SQl = "select b.* from meetingNotes1 a, loanApp b where a.appId=b.uid and a.rpId="&id&" order by b.memNo"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if not bdrs.eof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="700">
	<tr valign="top">
		<td height="25">否決下列貸款申請：</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>社員編號</td>
					<td>姓名</td>
					<td>貸款金額</td>
					<td>期數</td>
					<td>貸款理由摘要</td>
					<td>擔保人</td>
					<td>否決理由</td>
				</tr>
<%
do while not bdrs.eof%>
				<tr>
					<td><%=bdrs("memNo")%></td>
					<td><%=bdrs("memName")%></td>
					<td align="right"><%=formatnumber(bdrs("loanAmt"),2)%></td>
					<td><%=bdrs("installment")%></td>
<%
set reasonRs = conn.execute("select reasonName from reason a,loanReason b where a.uid=b.reasonID and b.loanAppID="&bdrs("uid"))
if not reasonRs.eof then
	thisReason = reasonRs.getString(,,,",")&bdrs("otherReason1")&bdrs("otherReason2")
	if right(thisReason,1)="," then thisReason=left(thisReason,len(thisReason)-1) end if
end if
%>
					<td><%if thisReason="" then response.write "&nbsp;" else response.write thisReason end if%></td>
					<td><%if isnull(bdrs("guarantorID")) then response.write "&nbsp;" else response.write bdrs("guarantorID") end if%></td>
					<td><%if isnull(bdrs("rejectReason")) then response.write "&nbsp;" else response.write bdrs("rejectReason") end if%></td>
				</tr>
<%
	bdrs.movenext
loop
%>
			</table>
		</td>
	</tr>
</table>
<br>
<%
end if

SQl = "select * from meetingNotes2 where rpId="&id&" order by memNo"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if not bdrs.eof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="700">
	<tr valign="top">
		<td height="25">批準下列借款人，簽字擔保者及擔保人撤回資金：</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>社員編號</td>
					<td>姓名</td>
					<td>撤回之金額</td>
					<td>摘要</td>
				</tr>
<%
do while not bdrs.eof%>
				<tr>
					<td><%=bdrs("memNo")%></td>
					<td><%=bdrs("memName")%></td>
					<td align="right"><%=formatnumber(bdrs("amount"),2)%></td>
					<td><%if bdrs("description")="" then response.write "&nbsp;" else response.write bdrs("description") end if%></td>
				</tr>
<%
	bdrs.movenext
loop
%>
			</table>
		</td>
	</tr>
</table>
<br>
<%
end if

SQl = "select * from meetingNotes3 where rpId="&id&" order by memNo"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if not bdrs.eof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="700">
	<tr valign="top">
		<td height="25">批準放棄下列期滿之副抵押品：</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>社員編號</td>
					<td>姓名</td>
					<td>貸款結餘</td>
					<td>期滿之副抵押品</td>
				</tr>
<%
do while not bdrs.eof%>
				<tr>
					<td><%=bdrs("memNo")%></td>
					<td><%=bdrs("memName")%></td>
					<td align="right"><%=formatnumber(bdrs("amount"),2)%></td>
					<td><%if bdrs("description")="" then response.write "&nbsp;" else response.write bdrs("description") end if%></td>
				</tr>
<%
	bdrs.movenext
loop
%>
			</table>
		</td>
	</tr>
</table>
<br>
<%
end if

SQl = "select * from meetingNotes4 where rpId="&id&" order by memNo"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if not bdrs.eof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="700">
	<tr valign="top">
		<td height="25">批準下列之延期合約：</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>社員編號</td>
					<td>姓名</td>
					<td>金額</td>
					<td>理由</td>
				</tr>
<%
do while not bdrs.eof%>
				<tr>
					<td><%=bdrs("memNo")%></td>
					<td><%=bdrs("memName")%></td>
					<td align="right"><%=formatnumber(bdrs("amount"),2)%></td>
					<td><%if bdrs("description")="" then response.write "&nbsp;" else response.write bdrs("description") end if%></td>
				</tr>
<%
	bdrs.movenext
loop
%>
			</table>
		</td>
	</tr>
</table>
<br>
<%
end if

SQl = "select * from meetingNotes5 where rpId="&id&" order by memNo"
Set bdrs = Server.CreateObject("ADODB.Recordset")
bdrs.open sql, conn, 3
if not bdrs.eof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="700">
	<tr valign="top">
		<td height="25">否決下列之延期合約：</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>社員編號</td>
					<td>姓名</td>
					<td>金額</td>
					<td>理由</td>
				</tr>
<%
do while not bdrs.eof%>
				<tr>
					<td><%=bdrs("memNo")%></td>
					<td><%=bdrs("memName")%></td>
					<td align="right"><%=formatnumber(bdrs("amount"),2)%></td>
					<td><%if bdrs("description")="" then response.write "&nbsp;" else response.write bdrs("description") end if%></td>
				</tr>
<%
	bdrs.movenext
loop
%>
			</table>
		</td>
	</tr>
</table>
<br>
<%
end if
otherAction = rs("otherAction")
if len(otherAction)>0 then
%>
<table border="0" cellpadding="0" cellspacing="0" width="700">
	<tr>
		<td height="25">除上列工作外，委員會亦曾採取下列行動：</td>
	</tr>
	<tr>
		<td><%=otherAction%></td>
	</tr>
</table>
<%
end if
%>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>