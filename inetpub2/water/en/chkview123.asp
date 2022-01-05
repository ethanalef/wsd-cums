.<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
Function clrTran(ByVal i )
   for i = 1  to 500
       store "" to acct(i)
       store "" to sdate(i)
       store "" to scode(i)
       store 0  to  shamt(i)
       store 0  to sbal(i)
       store "" to lcode(i)
       store 0  to lnpamt(i)
       store 0  to lniamt(i)
       store 0  to lnbal(i)
    next
    
end function
Sub ShT( xx ) 
         acct(xx)   =  ms("memno")
         sldate(xx) =  ms("ldate") 
         scode(xx)  =  ms("code")
         shamt(xx)  =  ms("amount")
         sbal(xx)   =  sbal(xx) + ms("amount")           
end Sub
Sub LnT(xx)
    lcode(xx)  = ms("code")
    lnpamt(xx) = ms("amount")
    lbal(xx)   = lbal(xx) + ms("amount")
end sub
server.scripttimeout = 1800	
y=year(now)
m=month(now)
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

xxx = yy+mm
randomize
xx = ROUND(rnd(xxx)*2000,0)
idx=round(rnd(XX)*26+1,0)
dim acct(500)
dim sdate(500)
dim scode(500)
dim shamt(500)
dim sbal(500)
dim lcode(500)
dim lnpamt(500)
dim lniamt(500)
dim lnbal(500)

xidx = "#temp"&idx
if request("submit")<>"" then
  conn.begintrans
  conn.execute( "create table "&xidx&"  ( memno int ,lnnum char(10), ldate smalldatetime, code char(2) , ammt money ) ")
  conn.execute("insert into   "&xidx&" (memno,ldate,code,ammt ) select a.memno,a.ldate,a.code,a.amount "&_
               "from share a inner join memmaster b on a.memno=b.memno "&_
               "where b.mstatus not in ('B','P','D','C') and b.accode='1950' "&_
               " order by a.memno,a.ldate,a.code ")
 conn.execute("insert into   "&xidx&" (memno,ldate,code,ammt ) select a.memno,a.ldate,a.code,a.amount "&_
               "from loan a inner join memmaster b on a.memno=b.memno "&_
               "where b.mstatus not in ('B','P','D','C') and b.accode='1950' "&_
               " order by a.memno,a.ldate,a.code ")
 set ms = conn.execute("select * from "&xidx&" order by memno,ldate,right(code,1),left(code,1)  ")
 if not ms.eof then
    xmemno=ms("memno")
    xldae =ms("ldate")
    xcode =ms("code")
    clrtran(500)
    xx = 0
    mx = 0
    do while not ms.eof
       if xmemno <> ms("memno") then
          response.write(xmemno)
       end if
       select case left(ms("code"),1)
              case "0"
                    call sht(xx)
                    ms.movenext
                    if ms("code") = "0D" then
                       call lnT(xx)
                   else
                       ms.moveepreviouse
                   end if
                   xx = xx + 1
              case "D"
                    select case ms("code")         
                           case "D0"                     
                                 acct(xx)   =  ms("memno")
                                 sldate(xx) =  ms("ldate") 
                                 lcode(xx)  = "貸款清數"
                                 lnpamt(xx) = ms("amount")
                                 lbal(xx)   = lbal(xx) - ms("amount")   
                             
 
                           case "D1"

                                 acct(xx)   =  ms("memno")
                                 sldate(xx) =  ms("ldate") 
                                 lnpamt(xx) = ms("amount")
                                 lbal(xx)   = ms("amount")
                                 set rs = conn.execute("select * from loanarec where lnnum='"&ms("lnnum")&"' ")
                                 if not rs.eof then
                                    if rs("lnflag") = "Y" then
 
                                       if rs("loantype")="E"  then 
                                          lniamt(xx)=""
                                          lnpamt(xx)=rs("appamt")
                                          lcode(xx)="*更改期數*"
                                       else
                                          lniamt(xx)=rs("chequeamt")&"+"
                                          lnpamt(xx)=rs("appamt")
                                          lcode(xx)="新貸款"                                      
                                       end if
 
                                 end if
                          end select                                 
                         
                    xx = xx + 1
              CASE "A"
                   select case ms("code")
                          case "A1"
                          case "A2"
                          case "A3"
                    ens select 
              case "B"

              cae  "C"
 
                                
              case "D"     
         
              case "E"
              case "F"
           end select  
                 
 
       ms.movenext
    loop
 end if
 ms.close
	conn.committrans
'	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>view contact</title>

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
<h3>wdate closed </h3>
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
