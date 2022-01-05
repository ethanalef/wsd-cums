<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
mFrom = request("mFrom")
mTo = request("mTo")
mMonth = cint(left(request("mPeriod"),2))
mYear = cint(right(request("mPeriod"),4))
mDate = mYear&"/"&mMonth&"/"&day(dateAdd("m",1,mYear&"/"&mMonth&"/1")-1)

SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close

if mMonth>8 then
	m=mMonth-8
	if mYear=acYear then y="this" else y="last" end if
else
	m=mMonth+4
	if mYear=acYear then y="last" else y="this" end if
end if

if m=1 then
	if y="this" then
		lastMonth = "lastBal12"
	else
		lastMonth = "openBal"
	end if
else
	lastMonth =y&"Bal"&m-1
end if

sql = "select glId,glName,glType,"&lastMonth&" as lastBal,"&y&"Bal"&m&" as thisBal from glMaster where deleted=0 and glId between '"&mFrom&"' and '"&mTo&"' and creationDate<='"&mDate&"' order by glId"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
elseif request.form("output")="text" then
	spaces=""
	spaces=space(50)
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Server.MapPath("..\txt")&"\"&session("username")&".txt", True)
	objFile.Write "WSDS Credit Union"
	objFile.WriteLine ""
	objFile.Write "G/L Monthly Transaction List"
	objFile.WriteLine ""
	objFile.Write left("Date"&spaces,10)
	objFile.Write left("Tran. No."&spaces,50)
	objFile.Write left("Particulars"&spaces,50)
	objFile.Write right(spaces&"Debit",20)
	objFile.Write right(spaces&"Credit",20)
	objFile.WriteLine ""
	for idx = 1 to 150
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		objFile.Write left(rs("glId")&" "&rs("glName")&spaces&spaces,92)
		objFile.Write "Beginning Balance:"
		if rs("thisBal")>0 then
			objFile.Write right(spaces&formatnumber(rs("thisBal"),2),20)
		else
			objFile.Write left(spaces,20)
		end if
		if rs("thisBal")<0 then
			objFile.Write right(spaces&formatnumber(-1*rs("thisBal"),2),20)
		else
			objFile.Write left(spaces,20)
		end if
		objFile.WriteLine ""
		sql = "select * from glTx where glId='" & rs("glId") & "' and deleted=0 and month(txDate)="&mMonth&" and year(txDate)="&mYear
		Set glRs = Server.CreateObject("ADODB.Recordset")
		glRs.open sql, conn
		do while not glRs.eof
			objFile.Write left(glRs("txDate")&spaces,10)
			objFile.Write left(" "&glRs("glTxNo")&spaces,50)
			objFile.Write left(glRs("txItem")&spaces,50)
			if glRs("txType")="D" then
				objFile.Write right(spaces&formatnumber(glRs("txAmt"),2),20)
			else
				objFile.Write left(spaces,20)
			end if
			if glRs("txType")="C" then
				objFile.Write right(spaces&formatnumber(glRs("txAmt"),2),20)
			else
				objFile.Write left(spaces,20)
			end if
			objFile.WriteLine ""
			glRs.movenext
		loop
		objFile.Write right(spaces&spaces&"Closing Balance:",110)
		if rs("lastBal")>0 then
			objFile.Write right(spaces&formatnumber(rs("lastBal"),2),20)
		else
			objFile.Write left(spaces,20)
		end if
		if rs("lastBal")<0 then
			objFile.Write right(spaces&formatnumber(-1*rs("lastBal"),2),20)
		else
			objFile.Write left(spaces,20)
		end if
		objFile.WriteLine ""
		for idx = 1 to 150
			objFile.Write "-"
		next
		objFile.WriteLine ""
		rs.movenext
	loop
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
<title>G/L Monthly Transaction List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="95%">
	<tr height="30" valign="top" align="center">
		<td colspan="5"><font size="4">WSDS Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="5"><font size="4">G/L Monthly Transaction List</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td><b>Date</b></td>
		<td><b>Tran. No.</b></td>
		<td><b>Particulars</b></td>
		<td align=right><b>Debit</b></td>
		<td align=right><b>Credit</b></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>
<%
do while not rs.eof
%>
	<tr height="15" valign="bottom">
		<td colspan=2><b><%=rs("glId")&" "&rs("glName")%></b></td>
		<td align=right><b>Beginning Balance:</b></td>
		<td align=right><b><%if rs("thisBal")>0 then response.write formatnumber(rs("thisBal"),2) end if%></b></td>
		<td align=right><b><%if rs("thisBal")<0 then response.write formatnumber(-1*rs("thisBal"),2) end if%></b></td>
	</tr>
<%
	SQl = "select * from glTx where glId='" & rs("glId") & "' and deleted=0 and month(txDate)="&mMonth&" and year(txDate)="&mYear
	Set glRs = Server.CreateObject("ADODB.Recordset")
	glRs.open sql, conn
	do while not glRs.eof
%>
	<tr height="15" valign="bottom">
		<td><%=glRs("txDate")%></td>
		<td><%=glRs("glTxNo")%></td>
		<td><%=glRs("txItem")%></td>
		<td align=right><%if glRs("txType")="D" then response.write formatnumber(glRs("txAmt"),2) end if%></td>
		<td align=right><%if glRs("txType")="C" then response.write formatnumber(glRs("txAmt"),2) end if%></td>
	</tr>
<%
		glRs.movenext
	loop
%>
	<tr height="15" valign="bottom">
		<td colspan=2></td>
		<td align=right><b>Closing Balance:</b></td>
		<td align=right><b><%if rs("lastBal")>0 then response.write formatnumber(rs("lastBal"),2) end if%></b></td>
		<td align=right><b><%if rs("lastBal")<0 then response.write formatnumber(-1*rs("lastBal"),2) end if%></b></td>
	</tr>
<%
	response.write "<tr><td colspan=5><hr></td></tr>"
	rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</center>
</body>
</html>
