<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
y=year(now)
m=month(now)-1
m = 2 
if m = 0 then
   m = 12
   y = y - 1
end if

server.scripttimeout = 1800

if request("submit")<>"" then
	mDate = y&"/"&m&"/"&request("mDay")
        if y/100=int(y/100) and y/4=int(y/4) then
           daylist ="312931303130313130313031"
           nday = mid(daylist,(m-1)*2+1,2)
        else
           daylist ="312831303130313130313031"
           nday = mid(daylist,(m-1)*2+1,2)         
	end if
                  

        set rs = server.createobject("ADODB.Recordset")
        set rs1 = server.createobject("ADODB.Recordset")
        set rs2 = server.createobject("ADODB.Recordset") 
        conn.begintrans

	sql = "select a.memno,a.adate,a.lnnum,a.code,a.bankin,a.curamt,a.updamt,a.status,a.flag,a.pdate,a.delyflag,a.mstatus,b.lnnum as xlnnum ,b.lndate,b.chkmon,b.months  from autopay a,loanrec b  where  right(a.code,1)='1' and  a.memno=b.memno and b.repaystat='N' union select memno,adate,lnnum,code,bankin,curamt,updamt,status,flag,pdate,delyflag,mstatus,null ,null,0,0 from autopay  where  right(code,1)='1'     and memno not in (select memno from loanrec  where repaystat='N')     "
 	rs.open sql, conn,1,1     

        do while not rs.eof
           memno = rs("memno")
           mm = month(rs("adate"))
           yy = year(rs("adate"))
           adate =dateSerial(yy,mm,20)
           lndate = rs("lndate")
           chkmon = rs("chkmon") - 1
           if rs("delyflag") = "Y" then
                conn.execute("update loanrec set chkmon = "&chkmon&"   where lnnum =  '"&rs("lnnum")&"' " )
                conn.execute("update loanrec set delyflag = 'N'   where lnnum =  '"&rs("lnnum")&"' and chkmon=0 " )

           end if 
           if rs("lnnum") = rs("xlnnum") then
              xlnnum = rs("lnnum")
            
           else
           if isNull(rs("xlnnum")) then
               xlnnum = rs("lnnum")
               
           else
              xlnnum = rs("xlnnum")
               
           end if
           end if 
          
           select case rs("flag")
                  case "F"
                       select case rs("code")
                              case "A1"                                                                      
                                      conn.execute("insert into share  (memno,ldate,code,amount,pflag,bal) values ('"&rs("memno")&"' ,'"&rs("adate")&"','AI',"&rs("bankin")&",0,"&rs("curamt")&" ) ")                                                                
                              case "E1"
                                      set ms=conn.execute("select * from loan where lnnum='"&rs("lnnum")&"' and ldate>='"&pdate&"' and code='E3'  ")
                                      if not ms.eof then
                                         conn.execute("insert into loan  (memno,lnnum,ldate,code,amount,pflag,bal) values ('"&rs("memno")&"','"&xlnnum&"' ,'"&rs("adate")&"','DE',"&rs("bankin")&",0,"&rs("curamt")&" ) ")                              
                                      else
                                         conn.execute("insert into loan  (memno,lnnum,ldate,code,amount,pflag,bal) values ('"&rs("memno")&"','"&xlnnum&"' ,'"&rs("adate")&"','DE',"&rs("bankin")&",1,"&rs("curamt")&" ) ")                              
                                      end if
                                      ms.close
                              case "F1"                    
                                      set ms=conn.execute("select * from loan where lnnum='"&rs("lnnum")&"' and ldate>='"&pdate&"' and code='F3'  ")
                                      if not ms.eof then
                                         conn.execute("insert into loan  (memno,lnnum,ldate,code,amount,pflag,bal) values ('"&rs("memno")&"','"&xlnnum&"' ,'"&rs("adate")&"','DF',"&rs("bankin")&",0,"&rs("curamt")&" ) ")                              
                                      else
                                         conn.execute("insert into loan  (memno,lnnum,ldate,code,amount,pflag,bal) values ('"&rs("memno")&"','"&xlnnum&"' ,'"&rs("adate")&"','DF',"&rs("bankin")&",1,"&rs("curamt")&" ) ")                              
                                      end if
                                      ms.close
                       end select

                  case else
                       select case rs("code")
                              case "A1"
  

                                    conn.execute("insert into share  (memno,ldate,code,amount) values ('"&rs("memno")&"' ,'"&rs("adate")&"','A1',"&rs("bankin")&" ) ")                              
                                       
                              case "E1"
 
                                   
                                       xx = rs("bankin") 
                                       set ms = conn.execute("select * from loan where memno='"&memno&"'  and  code='DE' and pflag=1 ")
                                       do while not ms.eof and xx > 0
                                          if xx >= ms("bal") then
                                             xx = xx - ms("bal")
                                             conn.execute("update loan set pflag=0 where memno='"&ms("memno")&"' and pflag=1 and code='DE' ")
                                          else
                                             yy = ms("bal") - xx
                                             xx = 0
                                             conn.execute("update loan set bal = "&yy&" where memno='"&ms("memno")&"' and pflag=1 and code='DE' ") 
                                          end if
                                         
                                        ms.movenext
                                        loop
                                        ms.close
                                      
                                      conn.execute("insert into loan  (memno,lnnum,ldate,code,amount) values ('"&rs("memno")&"','"&xlnnum&"' ,'"&rs("adate")&"','E1',"&rs("bankin")&" ) ")   
                                              
                                      conn.execute("update loanrec set bal=bal - "&rs("bankin")&" where  lnnum='"&xlnnum&"' ")
                                      conn.execute("update loanrec set cleardate ='"&rs("adate")&"' where lnnum='"&xlnnum&"'  and bal <=0 ")                                                              
                                      conn.execute("update loanrec set  repaystat= 'C'  where lnnum ='"&xlnnum&"'  and bal =0 ")                             
                                        conn.execute("update loanrec set  repaystat= 'N'  where lnnum ='"&xlnnum&"'  and bal <0 ")                             
                               
                              case "F1"

                                   
                                       xx = rs("bankin") 
                                       set ms = conn.execute("select * from loan where memno='"&memno&"' and  code='DF' and pflag=1 ")
                                       do while not ms.eof  and xx > 0
                                          if xx >= ms("bal") then
                                             xx = xx - ms("bal")
                                             conn.execute("update loan set pflag=0 where memno='"&ms("memno")&"' and pflag=1 and code='DF' ")
                                          else
                                             yy = ms("bal") - xx
                                             xx = 0
                                             conn.execute("update loan set bal = "&yy&" where memno='"&ms("memno")&"' and pflag=1 and code='DF' ") 
                                          end if
                                          
                                        ms.movenext
                                        loop
                                        ms.close
                                    
                                      conn.execute("insert into loan  (memno,lnnum,ldate,code,amount) values ('"&rs("memno")&"','"&xlnnum&"' ,'"&rs("adate")&"','F1',"&rs("bankin")&" ) ")   
                       end select
           end select
           if rs("status")="F" and rs("mstatus")<>"F"  then
              conn.execute("update memmaster set mstatus='"&rs("mstatus")&"' where memno='"&rs("memno")&"' ")
           end if
        rs.movenext
	loop


  
	rs.close
        conn.execute("update autopay set deleted= 1  where deleted=0 and right(code,1)='1' ")
	conn.committrans
       

        msg = "銀行轉帳過數巳完成!"
      response.redirect "completed.asp"
else
      id = ""
      set rs=conn.execute("select * from autopay where deleted = 0  and right(code,1)='1'  ")
       if rs.eof then
           msg = "銀行轉帳過數巳完成"
           id = "1"
        end if
        rs.close
end if
%>
<html>
<head>
<title>銀行轉帳過數</title>

<script language="JavaScript">
<!--


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;



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
<center>
<h3>銀行轉帳過數</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>

		<td width="10"></td>
		<td>

			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
