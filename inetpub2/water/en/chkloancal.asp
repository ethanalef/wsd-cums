<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
server.scripttimeout = 1800	
y=year(now)
m=month(now)

if request("submit")<>"" then

        lnnum = request("lnnum")        
       
        set rs = server.createobject("ADODB.Recordset")
        conn.begintrans
	
	sql = "select a.*,b.appamt from loan a,loanrec b where a.lnnum=b.lnnum and a.lnnum='"&lnnum&"'  order by a.memno,a.ldate,a.code "
	bal = 0
	rs.open sql, conn, 1, 1   
        if not rs.eof then
         xx = 0
         xbal = rs("appamt")    
        do while not rs.eof 
          
           select case rs("code")
                  case "E1","E2","E3"
                         monpay =  rs("amount")
                  case "F1","F2"
                      if xx > 0  then 
                      response.write(xbal)
                      response.write("<br>")
             
                       xint = xbal *.01
                       bal = bal + xint - rs("amount")
                   end if
                      xbal = xbal - monpay
                      xx = xx + 1
                    case "F3"
                         xbal = xbal - monpay
             end select
 
      
        rs.movenext
	loop
         end if
	rs.close
         
	conn.committrans
'	response.redirect "completed.asp"


end if
%>
<html>
<head>
<title>社員貸款利息差額連算</title>

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
		reqField=reqField+", ";
		if (!placeFocus)
			placeFocus=formObj.mDay;
	}

    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "J"+reqField.substring(2);
        else
	        reqField = "J"+reqField.substring(2,reqField.lastIndexOf(","))+''+reqField.substring(reqField.lastIndexOf(",")+2);
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.lnnum.focus()">

<br>
<center>
<h3>社員貸款利息差額連算</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
              <tr>
               	<td align="right" class="b12">貸款編號</td>
		<td width="10"></td>
		<td><input type="text" name="lnnum" value="<%=lnnum%>" size="11" >

                </td>
		   

        </tr>
<%if bal<>"" then%>
       <tr>
               	<td align="right" class="b12">差額</td>
		<td width="10"></td>
		<td ><%=bal%></td>
		   

        </tr>
<%end if%>
	<tr>
		<td>
			<input type="submit" value="計算" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>

