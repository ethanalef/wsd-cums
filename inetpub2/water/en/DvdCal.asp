<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%

server.scripttimeout = 1800


if request("submit")<>"" then
	conn.begintrans
	conn.execute("delete dividend")
	conn.committrans
	conn.begintrans
	myear  = request.form("myear")
	mrate = request.form("mrate")
	myr  = myear
	xxdate =dateserial(myr-1,7,1)
	yydate =dateserial(myr,7,1)
	ndate=dateserial(2008,4,30)

	if xxdate < ndate then
		xxdate = ndate
		chkdate ="30/04/"&(yy-1)
	end if

	dim sbal(50)
	dim divd(12)	' balance for dividend calculation
	dim lastBal(12)	' last balance for each dividend calculation month
	dim leastBal(12) ' least balance for each month
	dim dvdate(12)
	for i = 0 to 12
		' dividend calculation start from Jul 1 to End of Jun
		dvdate(i)=dateserial(myr-1, 7+i, 1)
	next

	gttlamt = 0
	subttl  = 0
	SQl = "select memno,memname,memcname,mstatus  from memmaster where   mstatus not in ('8','9','C','P','L','B' ) order by memno   "
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set ms = Server.CreateObject("ADODB.Recordset")
	Set ns = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn,1,1

	do while not rs.eof
	xmemno = rs("memno")
	mstatus = rs("mstatus")
	xx = 1
	sqlstr = "select * from share where memno='"&xmemno&"' and ldate<'"&yydate&"'  order by memno,ldate,code "
	ms.open sqlstr, conn,2,2
		if not ms.eof then
			for i = 1 to 50
				sbal(i) = 0
			next
			for i = 0 to 12
				leastBal(i) = -1
				lastBal(i) = -1
			next

			yy = 0
			do while not ms.eof
				if ms("ldate") > xxdate then
					' balance before Jul 1
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
					if left(ms("code"),1)="G" or left(ms("code"),1)="H" or left(ms("code"),1)="B" or  ms("code")="MF" then
						sbal(xx)=sbal(xx)-ms("amount")
					else
						if ms("code")="0A" or left(ms("code"),1)="C" or left(ms("code"),1)="A" then
							if ms("code")<>"AI" then
								sbal(xx) = sbal(xx) +ms("amount")
							end if
						end if
					end if
				end if
				ms.movenext
			loop

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

				if xmemno = 837 then
					aaaaaa = 0
				end if

				' balance of current dividend day = min(last balance of prev month, least balance of future month)
				subttl = 0
				for n = 0 to 11
					divd(n) = lastBal(n)
					for m = n + 1 to 12
						if divd(n) > leastBal(m) then
							divd(n) = leastBal(m)
						end if
					next
					subttl = subttl + formatNumber( divd(n) * mrate/100/12, 2) * 1
				next



			if mstatus="T" or mstatus="M" then
				conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'C' ,0) ")
			else
				if subttl <= 100 then
					conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'S',0 ) ")
				else
					conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'B',0 ) ")
				end if
			end if

			gttlamt = gttlamt + subttl

		end if
		ms.close

		rs.movenext
	loop

	conn.committrans
	check = 1
	msg="計算股息完成"
else
	mrate =2.00
	myear =year(date())
	check=0
end if

%>
<html>

<head>
<title>股息計算操作</title>

<script language="JavaScript">
<!--
function checkDay(mDay){
  D=mDay.value;
  M=<%=m%>;
  Y=<%=y%>;
  if(isNaN(D) || D=="")
    return false;
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
  return true;
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.myear.value==""){
		reqField=reqField+", 年份";
		if (!placeFocus)
			placeFocus=formObj.myear;
	}
	if (formObj.mrate.value==""){
		reqField=reqField+", 股息率";
		if (!placeFocus)
			placeFocus=formObj.mrate;
	}
    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "請填入"+reqField.substring(2);
        else
	        reqField = "請填入"+reqField.substring(2,reqField.lastIndexOf(","))+'及'+reqField.substring(reqField.lastIndexOf(",")+2);
        alert(reqField);
        placeFocus.focus();
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.myear.focus()">
<!-- #include file="menu.asp" -->
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<br>
<center>
<h3>股息計算操作</h3>
<form name="form1" method="post" action="dvdcal.asp">
<input type="hidden" name="check value="<%=check%>">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">年份</td>
		<td width="10"></td>
		<td><input type="text" name="myear" value="<%=myear%>" size="4" maxlength="4" >
         </tr>
         <tr>
                <td class="b8">股息率</td>
		<td width="10"></td>
		<td><input type="text" name="mrate" value="<%=mrate%>" size="5" maxlength="5" >
		<input type="submit" value="確定" name="submit" class="sbttn">
		</td>

	</tr>
<% if check=1 then %>
        <tr>
             <td class="b8">預計股息金額</td>
             <td width="10"></td>
	     <td id ="ttlamt"><%=formatnumber(gttlamt,2)%></td>
        </tr>
<%end if%>
</table>
</form>
</center>
</body>
</html>
