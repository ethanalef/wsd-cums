<!-- #include file="../conn.asp" -->
<%
if request("process")<>"" then
	conn.begintrans
	sql = "select * from glControl"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn, 2, 2
	acPeriod=rs("acPeriod")
	acYear=rs("acYear")
	rs.close
	set rs=nothing
	if acPeriod<=4 then
		m=acPeriod+8
		y=acYear
		oldestDate = y-1&"/9/1"
	else
		m=acPeriod-4
		y=acYear+1
		oldestDate = y-2&"/9/1"
	end if

	sql = "select a.*,b.txDate,b.amtLoan,b.monthlyRepaid from memMaster a Inner Join memTx b On a.memNo=b.memNo " &_
		"where a.deleted=0 and a.thisLoanBal"&acPeriod&">0 and b.amtLoan>0 Order By a.memNo, b.txDate Desc"
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.open sql, conn, 1, 3
	if not rs2.eof then
		thisMem = rs2("memNo")
		do while not rs2.eof
			if thisMen <> rs2("memNo") then
				thisMen = rs2("memNo")
				diff = DateDiff("m", cdate(year(rs2("txDate"))&"/"&month(rs2("txDate"))&"/1"), cdate(y&"/"&m&"/1"))

				response.write rs2("memNo") & " : " & rs2("amtLoan") & " : " & rs2("amtLoan")/rs2("monthlyRepaid") & " : " & rs2("OSinterest") & " : " & rs2("txDate") & " : "

				for idx = diff to 1 step -1
					if idx = diff then
						if month(rs2("txDate"))=12 then
							dayOfMonth = 31
						else
							dayOfMonth = day(cdate(year(rs2("txDate"))&"/"&month(rs2("txDate"))+1&"/1")-1)
						end if
						rs2("thisInterest") = (rs2("amtLoan")*0.01)-rs2("OSinterest")+int(((dayOfMonth-day(rs2("txDate")))/dayOfMonth*rs2("amtLoan")*0.01)+0.99)
					else
'						if idx-acPeriod<0 then
'							caly = "this"
'							calm = -1*(idx-acPeriod)
'						else
'							caly = "last"
'							calm = 12-(idx-acPeriod)
'						end if
						rs2("thisInterest") = (  (rs2("amtLoan") - ((diff-idx)*rs2("monthlyRepaid")))   *0.01)-rs2("OSinterest")+rs2("thisInterest")
					end if
				next
				rs2.update
				response.write dayOfMonth & " : " & rs2("thisInterest") & "<br>"
			end if
			rs2.movenext
		loop
	end if
	rs2.close
	set rs2=nothing
	conn.committrans
	conn.close
	set conn=nothing

	response.write "Done"
	response.end
end if
%>
<html>
<head>
<title>Re-Calculate Interest</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<br>
<center>
<h3>Re-Calculate Interest</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>">
<input type="submit" name="process" value="Process">
</form>
</center>
</body>
</html>
