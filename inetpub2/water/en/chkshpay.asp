<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
server.scripttimeout = 1800	
y=year(now)
m=month(now)

if request("submit")<>"" then
conn.begintrans
        
        set rs = server.createobject("ADODB.Recordset")
        set rs1= server.createobject("ADODB.Recordset")
        xdate1 = dateserial(2009,09,11)
        xdate2 = dateserial(2009,09,12)
        sql = "select * from shint order by memno "
        rs.open sql, conn, 1, 1   
        do while not rs.eof
           xmemno = rs("xmemno")
           shpay = rs("dividend_b")  
           shdiff = (rs("dividend_b") - rs("dividend_a"))
           if shdiff > 1  then
          sqlstr = "select * from dividend  where memno ='"&xmemno&"'  "
          rs1.open sqlstr, conn, 2, 2          
          if not  rs1.eof then
             rs1("dividend") = shpay
             rs1.update
           end if
           rs1.close
           xx = xx + 1
         end if
        rs.movenext
        loop
        rs.close 
        
conn.committrans
        response.write(xx)
        response.end
'	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>銀行轉賬自動過數</title>

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
