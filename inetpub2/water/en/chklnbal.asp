<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
server.scripttimeout = 1800	
y=year(now)
m=month(now)

if request("submit")<>"" then

        
        set rs = server.createobject("ADODB.Recordset")
        set rs1= server.createobject("ADODB.Recordset")
        conn.begintrans
	
	sql = "select a.memno,a.bal ,b.amount,A.LNNUM from loanrec a "&_
              "  inner join loan b on a.Lnnum=b.lnnum where a.REPAYstat ='N' and b.ldate>='2013/10/01' and (  code='E3') order by a.lnnum "
        

	rs.open sql, conn, 1, 1       
        do while not rs.eof 
           xbal = rs("bal") - rs(2)
         RESPONSE.WRITE(RS(0))
         RESPONSE.WRITE("===")
         RESPONSE.WRITE(RS(1))
         RESPONSE.WRITE("---")
         RESPONSE.WRITE(RS(1))
         RESPONSE.WRITE("---")       
         RESPONSE.WRITE(RS(2))
   
         RESPONSE.WRITE("***")
         RESPONSE.WRITE(XBAL)
         RESPONSE.WRITE("<BR>")

     conn.execute("UPDATE LOANREC SET BAL = '"&XBAL&"' WHERE LNNUM='"&RS("LNNUM")&"' ")
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
		<td class="b8">���</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=m%>/<%=y%>
			<input type="submit" value="�T�w" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>