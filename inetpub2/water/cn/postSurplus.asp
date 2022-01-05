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

	conn.begintrans
	sql = "select thisBal"&acPeriod&",glId from glMaster where deleted=0 and glType>=5"
	rs.open sql, conn, 2, 2
	ttl = 0
	do while not rs.eof
		ttl = ttl + round(rs(0),2)
		conn.execute("insert into glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'"&rs(1)&"','Surplus or Deficit','"&mDate&"','C',"&round(rs(0),2)&",0 from glTx")
		rs(0) = 0
		rs.update
		rs.movenext
	loop
	conn.execute("insert into glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0405','Surplus or Deficit','"&mDate&"','C',"&ttl&",0 from glTx")
	conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&ttl&" where glId='0405'")
	conn.committrans
	addUserLog "Post Surplus or Deficit"

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>歸納盈利</title>
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
<h3>歸納盈利</h3>
<%if acPeriod= 12 then%>
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
<%else%>
<p><font color="red">本部份只能於會計月份12月方能生效</font></p>
<%end if%>
</center>
</body>
</html>
