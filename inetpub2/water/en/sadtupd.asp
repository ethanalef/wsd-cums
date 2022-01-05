<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
	
y=year(now)
m=month(now)-1
if  m = 0 then
    m = 12
    y = y - 1
end if

if request("submit")<>"" then
	mDate = y&"/"&m&"/"&request("mDay")
        if y/100=int(y/100) and y/4=int(y/4) then
           daylist ="312931303130313130313031"
           nday = mid(daylist,(m-1)*2+1,2)
        else
           daylist ="312831303130313130313031"
           nday = mid(daylist,(m-1)*2+1,2)         
	end if
                  
        set rs1 = server.createobject("ADODB.Recordset")
        set rs = server.createobject("ADODB.Recordset")
         	conn.begintrans

	sql = "select * from autopay where right(code,1)='2' and deleted=0 "	
	rs.open sql, conn,1,1     
        do while not rs.eof
                 memno=rs("memno")
                 lnnum=rs("lnnum")
                 adate=rs("adate")
                 bankin=rs("bankin")

                 sql1 = "select * from loanrec where memno='"&rs("memno")&"' and repaystat='N' "
                 rs1.open sql1, conn,2,2
                 if not rs1.eof then
                     xlnnum=rs1("lnnum")
                 else
                    xlnnum=rs("lnnum")
                 end if
                 rs1.close
                  select case rs("code")
                        case "E2"
                             if rs("flag")="F" then
                                conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ( '"&rs("memno")&"','"&rs("lnnum")&"','"&rs("adate")&"','NE',"&rs("curamt")&") ")
                             else
                                if rs("lnnum")<>xlnnum then
                                 conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ( '"&rs("memno")&"','"&xlnnum&"','"&rs("adate")&"','E2',"&rs("bankin")&") ")
                                 set ms=conn.execute("select bal from loanrec where lnnum='"&xlnnum&"' ")
                                   if not ms.eof then
                                          xbal = ms(0) - rs("bankin")         
                                           conn.execute("update loanrec set bal= "&xbal&"  where lnnum='"&xlnnum&"' ")
                                    end if
                                   ms.close
                                 conn.execute("update loanrec set cleardate='"&rs("adate")&"'  where lnnum='"&xlnnum&"' and  bal=0 ")
                                 conn.execute("update loanrec set repaystat ='C'   where lnnum='"&xlnnum&"' and  bal='0' ") 
                                 conn.execute("update loanrec set repaystat ='N'   where lnnum='"&xlnnum&"' and  bal< '0' ") 
                                else
                                 conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ( '"&rs("memno")&"','"&rs("lnnum")&"','"&rs("adate")&"','E2',"&rs("bankin")&") ")
                                 set ms=conn.execute("select bal from loanrec where  lnnum='"&rs("lnnum")&"' ")
                                   if not ms.eof then
                                          xbal = ms(0) - rs("bankin")         
                                           conn.execute("update loanrec set bal= "&xbal&"  where lnnum='"&rs("lnnum")&"' ")
                                    end if
                                   ms.close
                                 conn.execute("update loanrec set cleardate='"&rs("adate")&"'  where lnnum='"&rs("lnnum")&"' and  bal=0 ")
                                 conn.execute("update loanrec set repaystat ='C'   where lnnum='"&rs("lnnum")&"' and  bal=0 ") 
                                end if  
                             end if
                            
                        case "F2"
                       
                             if rs("flag")="F" then
                                conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ( '"&rs("memno")&"','"&rs("lnnum")&"','"&rs("adate")&"','NF',"&rs("bankin")&") ")
                             else
                                if rs("lnnum") <> xlnnum then
                                    conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ( '"&rs("memno")&"','"&xlnnum&"','"&rs("adate")&"','F2',"&rs("bankin")&") ") 
                                else 
                                    conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ( '"&rs("memno")&"','"&rs("lnnum")&"','"&rs("adate")&"','F2',"&rs("bankin")&") ") 
                                end if 
                             end if 
                        case "A2"
                             if rs("flag")="F" then
                                conn.execute("insert into share (memno,ldate,code,amount) values ( '"&rs("memno")&"','"&rs("adate")&"','AI',"&rs("bankin")&") ")
                             else
                                 conn.execute("insert into share (memno,ldate,code,amount) values ( '"&rs("memno")&"','"&rs("adate")&"','A2',"&rs("bankin")&") ")
                             end if  
                end select    
        rs.movenext
	loop


rs.close
             conn.execute("update autopay set deleted= 1 where  deleted= 0 and right(code,1)='2'  ")
                msg = "庫房轉帳過數巳完成!"
	conn.committrans
	response.redirect "completed.asp"
else
      id = ""
      set rs=conn.execute("select * from autopay where deleted = 0  and right(code,1)='2'  ")
       if rs.eof then
           msg = "庫房轉帳過數巳完成"
           id = "1"
        end if
        rs.close
end if
%>
<html>
<head>
<title>庫房轉帳過數</title>

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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.submit.focus()">
<!-- #include file="menu.asp" -->
<br>

<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<center>
<h3庫房轉帳過數</h3>
<div><center><font size="3">庫房轉帳過數</font></center></div>
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<table border="0" cellpadding="0" cellspacing="0">
<% if id="" then %>
	<tr>

		<td width="10"></td>
		<td>

			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
<%end if%>
</table>
</form>
</center>
</body>
</html>
