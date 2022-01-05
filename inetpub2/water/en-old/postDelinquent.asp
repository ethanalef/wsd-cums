<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
sql = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
if acPeriod<=4 then
	m=acPeriod+8
	y=rs("acYear")
else
	m=acPeriod-4
	y=rs("acYear")+1
end if
rs.close

if request("submit")<>"" then
	mDate = y&"/"&m&"/"&request("mDay")

	set rs = conn.execute("select sum(round(thisLoanBal"&acPeriod&",2)) from memMaster where deleted=0 and overdue>=0 and overdue<=2")
	amt1to2 = rs(0)
	set rs = conn.execute("select sum(round(thisLoanBal"&acPeriod&",2)) from memMaster where deleted=0 and overdue>=3 and overdue<=6")
	amt3to6 = rs(0)
	set rs = conn.execute("select sum(round(thisLoanBal"&acPeriod&",2)) from memMaster where deleted=0 and overdue>=7 and overdue<=12")
	amt7to12 = rs(0)
	set rs = conn.execute("select sum(round(thisLoanBal"&acPeriod&",2)) from memMaster where deleted=0 and overdue>12")
	amt12 = rs(0)

	set rs = conn.execute("select thisBal"&acPeriod&" from glMaster where deleted=0 and glId='0201'")
	gl1to2 = rs(0)
	set rs = conn.execute("select thisBal"&acPeriod&" from glMaster where deleted=0 and glId='0202'")
	gl3to6 = rs(0)
	set rs = conn.execute("select thisBal"&acPeriod&" from glMaster where deleted=0 and glId='0203'")
	gl7to12 = rs(0)
	set rs = conn.execute("select thisBal"&acPeriod&" from glMaster where deleted=0 and glId='0204'")
	gl12 = rs(0)
	rs.close
	set rs=nothing

	if amt1to2+amt3to6+amt7to12+amt12=gl1to2+gl3to6+gl7to12+gl12 then
		conn.begintrans

		diff = amt3to6>gl3to6
		conn.execute("update glMaster set thisBal"&acPeriod&"="&amt3to6&" where glId='0202'")
		if diff>0 then
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0202','Delinquent','"&mDate&"','D',"&diff&",0 from glTx")
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delinquent','"&mDate&"','C',"&diff&",0 from glTx")
		else
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0202','Delinquent','"&mDate&"','C',"&-diff&",0 from glTx")
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delinquent','"&mDate&"','D',"&-diff&",0 from glTx")
		end if

		diff = amt7to12>gl7to12
		conn.execute("update glMaster set thisBal"&acPeriod&"="&amt7to12&" where glId='0203'")
		if diff>0 then
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0203','Delinquent','"&mDate&"','D',"&diff&",0 from glTx")
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delinquent','"&mDate&"','C',"&diff&",0 from glTx")
		else
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0203','Delinquent','"&mDate&"','C',"&-diff&",0 from glTx")
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delinquent','"&mDate&"','D',"&-diff&",0 from glTx")
		end if

		diff = amt12>gl12
		conn.execute("update glMaster set thisBal"&acPeriod&"="&amt12&" where glId='0204'")
		if diff>0 then
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0204','Delinquent','"&mDate&"','D',"&diff&",0 from glTx")
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delinquent','"&mDate&"','C',"&diff&",0 from glTx")
		else
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0204','Delinquent','"&mDate&"','C',"&-diff&",0 from glTx")
			conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Delinquent','"&mDate&"','D',"&-diff&",0 from glTx")
		end if

		conn.execute("update glMaster set thisBal"&acPeriod&"="&amt1to2&" where glId='0201'")
		conn.committrans
		addUserLog "Post Delinquent"

		conn.close
		set conn=nothing

		response.redirect "completed.asp"
	else
		msg = "呆賬不符總賬數目"
	end if
end if
conn.close
set conn=nothing
%>
<html>
<head>
<title>歸納呆賬</title>
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

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
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
<h3>歸納呆賬</h3>
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
<%if msg<>"" then%><p><font color="red"><%=msg%></font></p><%end if%>
</center>
</body>
</html>
