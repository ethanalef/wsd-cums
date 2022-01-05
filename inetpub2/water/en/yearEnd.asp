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
	mDividend = request("mDividend")

	conn.begintrans
	conn.execute("update memMaster set ttlLastShare=LastShrBal1 where deleted=0")
	for idx = 2 to 12
		conn.execute("update memMaster set ttlLastShare=ttlLastShare + FLOOR(LastShrBal"&idx&" / 5) * 5 where deleted=0")
		conn.execute("update memMaster set ttlLastShare=FLOOR(LastShrBal"&idx&" / 5) * 5 * "&idx&" where LastShrBal"&idx&"<LastShrBal"&idx-1&" and deleted=0")
	next
	set rs = conn.execute("select sum(round(ttlLastShare/12*"&mDividend&"/100,2)) from memMaster where ttlLastShare>5 and deleted=0")
	conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&rs(0)&" where glId='0205'")
	conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Dividend','"&mDate&"','C',"&rs(0)&",0 from glTx")
	conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&rs(0)&" where glId='0401'")
	conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0401','Dividend','"&mDate&"','D',"&rs(0)&",0 from glTx")
	conn.execute("update memMaster set dividend=round(ttlLastShare/12*"&mDividend&"/100,2), thisShrBal"&acPeriod&"=thisShrBal"&acPeriod&"+round(ttlLastShare/12*"&mDividend&"/100,2) where ttlLastShare>5 and deleted=0")
	conn.committrans

	addUserLog "Dividend process"

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>年結股息計算</title>
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

	if (formObj.mDividend.value==""||!formatNum(formObj.mDividend)){
		reqField=reqField+", 股息率";
		if (!placeFocus)
			placeFocus=formObj.mDividend;
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
<h3>年結股息計算</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">日期</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=m%>/<%=y%>
		</td>
	</tr>
	<tr>
		<td class="b8">股息率</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDividend" value="<%=mDividend%>" size="8" maxlength="8" onblur="if(!formatNum(this)){this.value=''};">%
			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
