<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
server.scripttimeout = 1800	
y=year(now)
m=month(now)

if request("submit")<>"" then

        
         set rs = server.createobject("ADODB.Recordset")
        conn.begintrans
	
	sql = "select * from loanrecold where memno='2322' order by memno,lndate "
	
	rs.open sql, conn, 1, 1       
        do while not rs.eof 
        conn.execute(" INSERT INTO Loanrec (memno ,lnnum,lndate ,appamt,install  ,monthrepay,cleardate ,repaystat ,bal ,lnflag ,chequeamt ,oldlnnum  ,calflag ,delyflag,months ,chkmon ,delydate  ,applyamt ,loantype ,amount) VALUES ( '"&rs("memno")&"', '"&rs("lnnum")&"', '"&ymd(rs("lndate"))&"' , '"&rs("appamt")&"', '"&rs("install")&"' , '"&rs("monthrepay")&"' , '"&ymd(rs("cleardate"))&"' , '"&rs("repaystat")&"', '"&rs("bal")&"' , '"&rs("lnflag")&"' , '"&rs("chequeamt")&"' ,'"&rs("oldlnnum")&"' , '"&rs("calflag")&"' ,'"&rs("delyflag")&"' ,'"&rs("months")&"'  ,'"&rs("chkmon")&"' , '"&ymd(rs("delydate"))&"'  ,'"&rs("applyamt")&"'  ,'"&rs("loantype")&"'  ,'"&rs("amount")&"'  ) ")
        rs.movenext
	loop
	rs.close

	conn.committrans
'	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>�Ȧ����۰ʹL��</title>

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
		reqField=reqField+", ���";
		if (!placeFocus)
			placeFocus=formObj.mDay;
	}

    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "�ж�J"+reqField.substring(2);
        else
	        reqField = "�ж�J"+reqField.substring(2,reqField.lastIndexOf(","))+'��'+reqField.substring(reqField.lastIndexOf(",")+2);
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
<h3>�Ȧ����۰ʹL��</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<input type="submit" value="�T�w" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>

