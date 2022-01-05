<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQl = "select a.*,b.memname from sadttran a,memmaster b where a.memno=b.memNo  order by a.memno"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

ttlPerm=0 : ttlTemp=0
A1ttl = 0
E1ttl = 0
F1ttl = 0
AIttl = 0
FIttl = 0
EIttl = 0

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
	objFile.Write "Water Supplies Department Staff Credit Union"
	objFile.WriteLine ""
	objFile.Write "A/C Check List for Salary Deduction"
	objFile.WriteLine ""	
	objFile.Write left("A/C No."&spaces,10)
	objFile.Write left("Name"&spaces,50)
	objFile.Write right(spaces&"(Code)",10)
	objFile.Write right(spaces&"(Amount)",20)
	objFile.WriteLine ""
	for idx = 1 to 100
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&spaces,50)
		objFile.Write right(spaces&rs("trefno"),10)
		objFile.Write right(spaces&formatnumber(rs("amount"),2),20)
		objFile.WriteLine ""
 	       if rs("trefno") = "A2" then a1ttl =  a1ttl + rs("amount") end if
       	       if rs("trefno") = "AI" then aIttl =  aIttl + rs("amount") end if
       	       if rs("trefno") = "E2" then E1ttl =  E1ttl + rs("amount") end if
      	       if rs("trefno") = "EI" then a1ttl =  EIttl + rs("amount") end if
               if rs("trefno") = "F2" then F1ttl =  F1ttl + rs("amount") end if
               if rs("trefno") = "FI" then DIttl =  DIttl + rs("amount") end if		
		ttlTemp=ttlTemp+round(rs("amount"),2)
		rs.movenext
	loop
	for idx = 1 to 100
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(70)
	objFile.Write right(spaces&formatnumber(ttlTemp,2),20)
	objFile.WriteLine ""

	objFile.Write left("Auto-Saving Amount"&spaces,40)        
	objFile.Write right(spaces&formatnumber(a1ttl,2),20)
	objFile.WriteLine ""
        if aittl <> 0 then
	objFile.Write left("Auto-Saving Lost Amount"&spaces,40)        
	objFile.Write right(spaces&formatnumber(aittl,2),20)
	objFile.WriteLine ""
        end if
	objFile.Write left("Auto-Loan Paid Amount"&spaces,40)        
	objFile.Write right(spaces&formatnumber(E1ttl,2),20)
	objFile.WriteLine ""
        if eittl <> 0 then
	objFile.Write left("Auto-Loan Paid Lost Amount"&spaces,40)        
	objFile.Write right(spaces&formatnumber(Eittl,2),20)
	objFile.WriteLine ""
        end if
	objFile.Write left("Auto-Interest paid Amount"&spaces,40)        
	objFile.Write right(spaces&formatnumber(F1ttl,2),20)
	objFile.WriteLine ""
        if fittl <> 0 then
	objFile.Write left("Auto-Interest Paid Lost Amount"&spaces,40)        
	objFile.Write right(spaces&formatnumber(Fittl,2),20)
	objFile.WriteLine ""
        end if
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
<title>A/C Check List for Auto-pay</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="4">Water Supplies Department Staff Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">A/C Check List for Salary Deduction</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>A/C No.</b></td>
		<td width="200"><b>Name</b></td>
		<td width="130" align="right"><b>(Code)</b></td>
		<td width="130" align="right"><b>Auto-pay<br>(Amount)</b></td>
	</tr>
	<tr><td colspan=4><hr></td></tr>
<%
do while not rs.eof
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td align="right"><%=rs("trefno")%></td>
		<td align="right"><%=formatNumber(rs("amount"),2)%></td>
	</tr>
<%
        if rs("trefno") = "A2" then a1ttl =  a1ttl + rs("amount") end if
       if rs("trefno") = "AI" then aIttl =  aIttl + rs("amount") end if
       if rs("trefno") = "E2" then E1ttl =  E1ttl + rs("amount") end if
       if rs("trefno") = "EI" then a1ttl =  EIttl + rs("amount") end if
       if rs("trefno") = "F2" then F1ttl =  F1ttl + rs("amount") end if
       if rs("trefno") = "FI" then DIttl =  DIttl + rs("amount") end if


	ttlTemp=ttlTemp+round(rs("amount"),2)
	rs.movenext
loop
%>
	<tr><td colspan=4><hr></td></tr>
	<tr>
		<td colspan="2"></td>
		<td></td>
		<td align="right"><%=formatNumber(ttlTemp,2)%></td>
	</tr>
	
	<tr>
		
		<td>Auto-Saving Amount</td>
		<td align="right"><%=formatNumber(a1ttl,2)%></td>
		<td colspan="2"></td>
		
	</tr>
<%    if ai1ttl <> 0 then %>
	<tr>
		<td>Auto-Saving Lost Amount</td>
		<td align="right"><%=formatNumber(aittl,2)%></td>
		<td colspan="2"></td>
	</tr>
<% end if %>
	<tr>
		
		<td>Auto-Loan Paid Amount</td>
		<td align="right"><%=formatNumber(e1ttl,2)%></td>
		<td colspan="2"></td>
		
	</tr>
<%    if eittl <> 0 then %>
	<tr>
		<td>Auto-Loan Paid Lost Amount</td>
		<td align="right"><%=formatNumber(eittl,2)%></td>
		<td colspan="2"></td>
	</tr>
<% end if %>
	<tr>
		
		<td>Auto-Interest Paid Amount</td>
		<td align="right"><%=formatNumber(F1ttl,2)%></td>
		<td colspan="2"></td>
		
	</tr>
<%   if fittl <> 0 then %>
	<tr>
		<td>Auto-Interest Paid Lost Amount</td>
		<td align="right"><%=formatNumber(Fittl,2)%></td>
		<td colspan="2"></td>
	</tr>
<% end if  %>
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
