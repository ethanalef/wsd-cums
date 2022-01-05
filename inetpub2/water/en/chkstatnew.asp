<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
	
y=year(now)
m=month(now)

if request("submit")<>"" then
           conn.begintrans 
        conn.execute("delete memmaster")
	conn.committrans
        set rs = server.createobject("ADODB.Recordset")
        conn.begintrans
	
	sql = "select * from chknew1 "
	
	rs.open sql, conn, 1, 1
        set rs1 = server.createobject("ADODB.Recordset") 
	sql1 = "select * from memmaster where 0=1 "
	
	rs1.open sql1, conn, 2, 2       
        do while not rs.eof 
 
           rs1.addnew
           rs1("memno") = rs("acct")
           rs1("memname") = rs("name")
           rs1("memcname") = rs("cname")
           rs1("memgrade") = rs("post")
           rs1("memaddr1") = rs("add1")
           rs1("memaddr2") = rs("add2")
           rs1("memcontacttel")=rs("telh")
           rs1("memofficetel")=rs("telo")  
           rs1("mstatus") = rs("status")
           rs1("bnk") = rs("bnk")
           rs1("bch")=rs("bch")
           rs1("bacct")=rs("bacct")
           rs1("bnklmt")=rs("bnklmt")
            rs1("memBday")= rs("dob")
           rs1("memDate")=rs("dtjoint") 
            rs1("wdate")=rs("dtclose")
           select case rs("status") 
                  case "T","M"  
                      rs1("monthssave") = rs("save")
                      rs1("monthsave")=0  
                      rs1("tpayamt")=rs("tryamt")
                      
                  case "A"

                      rs1("monthssave")= 0
                  if rs("saving") > 0  then
                      rs1("monthsave")=rs("saving")
                  else
                      rs1("monthsave")= rs("save")
                  end if 
                      rs1("tpayamt")=0 
           end select                  
           if rs("EMAIL") <> "x" then   
	       rs1("mememail")  = rs("EMAIL")	
           end if
	   rs1("memhkid") = rs("id")
           rs1("memgender") = rs("sex")
           rs1("membday") = rs("dob")	
          
              rs1("bnklmt")  = rs("bnklmt")
          
           rs1("deleted") = 0
           rs1.update
       
        rs.movenext
	loop
        rs1.close
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
