<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
server.scripttimeout = 1800	
y=year(now)
m=month(now)

if request("submit")<>"" then

        xdate = "2007.06.01"
        set rs = server.createobject("ADODB.Recordset")
        set rs1= server.createobject("ADODB.Recordset")

        conn.begintrans
       conn.execute("update loanrec set oldlnnum= null where  lnflag<>'Y' and repaystat='N' ")         
	sql =  "select * from loanrec where repaystat = 'N'    "
        rs.open sql, conn, 1, 1          
        do while not rs.eof
           if rs("lnflag") ="Y" then
                 yy = year(rs("lndate"))
                 mm = month(rs("lndate"))
                 dd = day(rs("lndate"))
                                
                if ((yy/4)= int(yy/4) and (yy/100)=int(yy/100)) then
                   daylist="312931303130313130313031"
                   mD = mid(daylist,(mm-1)*2+1,2)
                else
                   daylist="312831303130313130313031"
                   mD = mid(daylist,(mm-1)*2+1,2)
                end if                  
                   set rs1 = conn.execute("select amount from loan where lnnum ='"&rs("oldlnnum")&"' and code ='D0' ")
                   if not rs1.eof then
                      bal =rs("chequeamt") +  rs1(0)
                      int1 = rs1(0)*0.01
                      int2 = round(rs("chequeamt")*0.01*(md - dd+1)/md,2)
                      ttlint = int1 + int2
                    
                       
 
                   set rs2 = conn.execute("select  *  from loan where lnnum ='"&rs("lnnum")&"' and ldate>='"&rs("lndate")&"' order by lnnum,ldate,right(code,2),left(code,1)   ")
                   do while  not rs2.eof 
                      select case left(rs2("code"),1)
                             case "E"
                                   bal = bal - rs2("amount")
                      end select
                   rs2.movenext
                   loop                                    
                   rs2.close
                   rs1.close
                   conn.execute("update loanrec set bal="&bal&" where lnnum ='"&rs("lnnum")&"' and repaystat='N' ")
                  end if 
                 else
                  cnt = 0
                  xbal = rs("appamt") 
                  bal = xbal
                  set rs2 = conn.execute("select  *  from loan where lnnum ='"&rs("lnnum")&"' and ldate>='"&rs("lndate")&"' order by lnnum,ldate,right(code,2),left(code,1)   ")
                   do while  not rs2.eof 
                      select case left(rs2("code"),1)
                             case "E"
                                   xbal = xbal - rs2("amount")
                                   cnt = cnt + 1
                      end select
                   rs2.movenext
                   loop                                    
                   rs2.close
 
                   conn.execute("update loanrec set bal="&xbal&" where lnnum ='"&rs("lnnum")&"' and repaystat='N' ")                          
                
                  

           end if      


 
        rs.movenext
        loop
        rs.close
	conn.committrans
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
