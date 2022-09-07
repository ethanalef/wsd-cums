<!-- #include file="../conn.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="init.asp" -->

<%
	server.scripttimeout = 1800

	dmon = mid(request.form("dvdDay"), 4, 2)
	dday = left(request.form("dvdDay"), 2)
	rate = request.form("rate")

	yr =request.form("yr")
	yr1yr2 = (yr-1) & "/" & yr
%>

<html>
	<head>
		<title>股息全年結</title>
		<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
		<link href="../main.css" rel="stylesheet" type="text/css">
		<style type='text/css'>p {page-break-after: always;}</style>
	</head>
	<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
		<centre>
			<table border="0" cellpadding="0" cellspacing="0">
				<tr height="30" valign="top" align="center">
					<td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>股息全年結<br>
						<font size="2"  face="標楷體" ><%=yr1yr2%></font></font>
					</td>
				</tr>
				<tr height="30" >
					<td colspan=9></td>
				</tr>
				<tr height="15" valign="bottom">
					<td width="80" align="center"><font size="2"  face="標楷體" >社員編號</font></td>
					<td width="80"  align="center"><font size="2"  face="標楷體" >  姓名</font</td>
					<td width="130" align="right"><font size="2"  face="標楷體" > 金額</fot></td>
				</tr>
				<tr><td colspan=6><hr></td></tr>
<%
	xxdate = dateserial(yr-1, 7, 1)
	yydate = dateserial(yr, 7, 1)
	mmdate = (yr-1) & "/07/01"
	nndate = yr & "/07/01"
	chkdate ="01/07/" & (yy-1)
	ndate = dateserial(2008,4,30)
	if xxdate < ndate then
	xxdate = ndate
	chkdate ="30/04/" & (yy-1)
	end if

	dim sbal(50)
	dim divd(12)	' balance for dividend calculation
	dim lastBal(12)	' last balance for each dividend calculation month
	dim leastBal(12) ' least balance for each month
	dim dvdate(12)
	dim sdivd
	dim ttlDivd

	for i = 0 to 12
		' dividend calculation start from Jul 1 to End of Jun
		dvdate(i)=dateserial(yr-1, 7+i, 1)
	next

	sql = "select a.memno, a.memcname, mstatus from memmaster a "&_
			"where exists ("&_
				"select memno from share as b "&_
					"where ldate < '"&nndate&"' and ldate >= '"&mmdate&"' and a.memNo = b.memno and "&_
						"mstatus not in ('C','P','B','I' ) and wdate is null) "&_
			"order by a.memno"

	Set rs = Server.CreateObject("ADODB.Recordset")
	Set ms = Server.CreateObject("ADODB.Recordset")
	Set ns = Server.CreateObject("ADODB.Recordset")

	if request.form("output")="Word" then
		Response.ContentType = "application/msword"
	elseif request.form("output")="excel" then
		Response.ContentType = "application/vnd.ms-excel"
	end if

	rs.open sql, conn,1,1
	do while not rs.eof
		mno = rs("memno")
		mcname = rs("memcname")
		mstatus = rs("mstatus")

		set ms = conn.execute("select * from share where memno='"&mno&"' and ldate < '"&yydate&"'  order by memno,ldate,code ")
		if not ms.eof then
			for i = 1 to 50
				sbal(i) = 0
			next
			for i = 0 to 12
				leastBal(i) = -1
				lastBal(i) = -1
			next
			xx = 1
			yy = 0
			
			do while not ms.eof
				xyear = year(ms("ldate"))
				xmon  = month(ms("ldate"))
				xday  = day(ms("ldate"))
				xdate = xyear&xmon&xday
				ssdate = right("0"&xday,2)&"/"&right("0"&xmon,2)&"/"&xyear

				if ms("ldate") > xxdate then
					if sbal(xx) > 0 and  yy = 0 then
						leastBal(0) = sbal(xx)
						lastBal(0) = sbal(xx)
						xx = xx + 1
						yy = 1
					end if

					select case ms("code")
						case	"0A"
								sbal(xx)=sbal(xx-1)+ms("amount")
						case	"A8"
								sbal(xx)=sbal(xx-1)+ms("amount")
						case	"B8"
								sbal(xx)=sbal(xx-1)-ms("amount")
						case 	"B1","MF"
								sbal(xx)=sbal(xx-1)-ms("amount")
						case	"G3","H3"
								sbal(xx)=sbal(xx-1)-ms("amount")
						case 	"G0" ,"H0","B0","B3","BE","BF"
								sbal(xx)=sbal(xx-1)-ms("amount")
						case  	"AI","CH"
								sbal(xx) = sbal(xx-1)
						case 	"A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4"
								sbal(xx) = sbal(xx-1) + ms("amount")
					end select

					divCalDate = dateserial(year(ms("ldate")),month(ms("ldate"))+1, 1)
					for i = 0 to 12
						if dvdate(i) = divCalDate and ms("code") <> "A8" and ms("code") <> "B8" then
							if sbal(xx) < leastBal(i) and leastBal(i) > -1 or leastBal(i) = -1 then
								leastBal(i) = sbal(xx)
							end if
							lastBal(i) = sbal(xx)
							exit for
						end if
					next

					xx = xx + 1

				else
					' balance before Jul 1
					if left(ms("code"),1)="G" or left(ms("code"),1)="H" or left(ms("code"),1)="B" or ms("code")="MF" then
						sbal(xx)=sbal(xx)-ms("amount")
					else
						if ms("code")="0A" or left(ms("code"),1)="C" or left(ms("code"),1)="A" then
							if ms("code")<>"AI"  then
								sbal(xx) = sbal(xx) +ms("amount")
							end if
						end if
					end if
				end if
				ms.movenext
			loop
		end if
		ms.close

		' balance propagation
		if lastBal(0) = -1 then
			lastBal(0) = 0
			end if
		if leastBal(0) = -1 then
			leastBal(0) = 0
		end if
		for i = 1 to 12
			if leastBal(i) = -1 then
				leastBal(i) = leastBal(i-1)
			end if
			if lastBal(i) = -1 then
				lastBal(i) = lastBal(i-1)
			end if
		next

		sdivd = 0
		' balance of current dividend day = min(last balance of prev month, least balance of future month)
		for n = 0 to 11
			divd(n) = lastBal(n)
			for m = n + 1 to 12
				if divd(n) > leastBal(m) then
					divd(n) = leastBal(m)
				end if
			next
			sdivd = sdivd + formatnumber(divd(n) * rate / 100 /12, 2) * 1
		next

		if sdivd > 0 then
			ttlDivd = ttlDivd + sdivd
%>
				<tr>
						<td width="80" align="center"><%=mno%></td>
						<td width="80" align="center" ><font size="2"  face="標楷體" ><%=mcname%></font></td>
						<td width="130" align="right"><%=formatnumber(sdivd, 2)%></td>
				</tr>
<%
		end if

		rs.movenext
	loop

	rs.close
	set rs = nothing
	conn.close
	set conn = nothing
%>
				<tr><td colspan=6><hr></td></tr>
				<tr>
					<td></td>
          <td></td>             
          <td width="130" align="right"><%=formatnumber(ttlDivd, 2)%></td>
        </tr>
			</table>
		</centre>
	</body>
</html>

