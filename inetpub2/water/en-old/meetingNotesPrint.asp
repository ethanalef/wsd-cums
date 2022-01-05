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
		<td colspan="2" align="center" height="40"><font size="4">���q�u�{�p�x�W���U���U�کe���|�|ĳ�O��</font></td>
	</tr>
	<tr valign="top">
		<td colspan="2" height="90">
			�}�|����G <%=year(rs("rpDate"))%> �~ <%=month(rs("rpDate"))%> �� <%=day(rs("rpDate"))%> ��C���~�ײ� <%=rs("rpNo")%> ���|ĳ�C
			�� <%=pageno%> ��<br>
			�U�کe���| <%=rs("rpType")%> �b�W�C����� <%=rs("rpTime")%> �}�l�b������ƳB�|��C<br>
			�e�� <%=year(rs("lastRpDate"))%> �~ <%=month(rs("lastRpDate"))%> �� <%=day(rs("lastRpDate"))%>
			��|ĳ���|ĳ�O���gŪ�X��q�L�A�å[�H�U�C�ץ��G<br>
			<%=rs("amenment")%>
		</td>
	</tr>
	<tr valign="top" height="20">
		<td>�X�u�e�� �G<%=rs("present")%></td><td align="right" width="100">�ʮu�̡G<%=rs("absent")%></td>
	</tr>
	<tr valign="top" height="60">
		<td colspan="2">
			�X�u����L�H�K�G<%=rs("attendee")%><br>
			�����|ĳ�Ų߳��{�G<%=rs("overview")%><br>
			��ǤU�C�f�ڥӽСG<br>
		</td>
	</tr>
	<tr valign="top" height="520">
		<td colspan="2">
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>&nbsp;</td>
					<td>�����s��</td>
					<td>�m�W</td>
					<td>�U�ڪ��B</td>
					<td>����</td>
					<td>�U�ڲz�ѺK�n�Ω��</td>
					<td>��O�H</td>
					<td>�Ƶ�</td>
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
					<td>�����U�X���B</td><td align="center" width="160"><%=formatnumber(thisLoanAmt,2)%></td>
					<td>���A�@</td><td align="center" width="160"><%=thisMem%></td><td>�H</td>
				</tr>
				<tr>
					<td>����O���@</td><td align="center"><%=-int(-bdrs.recordcount/rowPerPage)%></td>
					<td>���A�`���B</td><td align="center"><%=formatnumber(totalLoanAmt,2)%></td><td>��</td>
				</tr>
				<tr>
					<td>�`�H��</td><td align="center"><%=bdrs.recordcount%></td>
					<td>�H�A�@����</td><td align="center"><%=rs("interview")%></td><td>�H</td>
				</tr>
				<tr>
					<td>���ɤ���</td><td align="center"><%=totalNo1%></td>
					<td>�H�A�U�X���B</td><td align="center"><%=formatnumber(totalLoanAmt1,2)%></td><td>��</td>
				</tr>
				<tr>
					<td>�Ͳ��γ~</td><td align="center"><%=totalNo2%></td>
					<td>�H�A�U�X���B</td><td align="center"><%=formatnumber(totalLoanAmt2,2)%></td><td>��</td>
				</tr>
				<tr>
					<td>�_�M�U�ڥӽ�</td><td align="center"><%=rejectNo%></td>
					<td>�v�A�@</td><td align="center"><%=formatnumber(rejectAmt,2)%></td><td>��</td>
				</tr>
			</table>
		<td>
	</tr>
	<tr>
		<td colspan="2">
		�U���|ĳ����G <%=year(rs("nextRpDate"))%> �~ <%=month(rs("nextRpDate"))%> �� <%=day(rs("nextRpDate"))%> ��A<%=rs("nextRpTime")%>�_������ƳB�|��C
		<br><br><br><br><br>
		ñ�p�G ________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ________________________ <br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �D�u (���è})
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���� (��ģ��)
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
		<td height="25">�_�M�U�C�U�ڥӽСG</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>�����s��</td>
					<td>�m�W</td>
					<td>�U�ڪ��B</td>
					<td>����</td>
					<td>�U�ڲz�ѺK�n</td>
					<td>��O�H</td>
					<td>�_�M�z��</td>
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
		<td height="25">��ǤU�C�ɴڤH�Añ�r��O�̤ξ�O�H�M�^����G</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>�����s��</td>
					<td>�m�W</td>
					<td>�M�^�����B</td>
					<td>�K�n</td>
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
		<td height="25">��ǩ��U�C�������Ʃ��~�G</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>�����s��</td>
					<td>�m�W</td>
					<td>�U�ڵ��l</td>
					<td>�������Ʃ��~</td>
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
		<td height="25">��ǤU�C�������X���G</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>�����s��</td>
					<td>�m�W</td>
					<td>���B</td>
					<td>�z��</td>
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
		<td height="25">�_�M�U�C�������X���G</td>
	</tr>
	<tr valign="top">
		<td>
			<table border="1" cellpadding="3" cellspacing="0">
				<tr align="center">
					<td>�����s��</td>
					<td>�m�W</td>
					<td>���B</td>
					<td>�z��</td>
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
		<td height="25">���W�C�u�@�~�A�e���|�紿�Ĩ��U�C��ʡG</td>
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