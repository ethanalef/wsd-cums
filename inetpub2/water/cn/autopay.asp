<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
sql = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close
if acPeriod<=4 then
	m=acPeriod+8
	y=acYear
else
	m=acPeriod-4
	y=acYear+1
end if

if request("submit")<>"" then
	mDate = y&"/"&m&"/"&request("mDay")

	conn.begintrans
	set rs = conn.execute("select sum(autopayAmt) from memMaster where autopayAmt<>0 and deleted=0")
	conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&rs(0)&" where glId='0205'")
	conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Auto-Pay','"&mDate&"','D',"&rs(0)&",0 from glTx")
	conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&rs(0)&" where glId='0401'")
	conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0401','Auto-Pay','"&mDate&"','C',"&rs(0)&",0 from glTx")
	conn.execute("update memMaster set thisShrBal"&acPeriod&"=thisShrBal"&acPeriod&" + autopayAmt where deleted=0")
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql = "select autopayAmt, memNo from memMaster where autopayAmt<>0 and deleted=0 order by memNo"
	rs2.open sql, conn, 3
	do while not rs2.eof
		conn.execute("insert into memTx (memTxNo,memNo,txDate,treNo,sharePaid,txAmt,deleted) select max(memTxNo)+1,'"&rs2(1)&"','"&mDate&"','AT',"&rs2(0)&","&rs2(0)&",0 from memTx")
		rs2.movenext
	loop
	rs2.close
	set rs2=nothing
	conn.committrans

	addUserLog "Autopay process"

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>銀行轉賬自動過數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
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

	if (formObj.mDay.value==""){
		reqField=reqField+", 日期";
		if (!placeFocus)
			placeFocus=formObj.mDay;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mDay.focus()">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>銀行轉賬自動過數</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">日期</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=m%>/<%=y%>
			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
