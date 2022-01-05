<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%

server.scripttimeout = 1800

 




if request("submit")<>"" then
 
   conn.begintrans   
   conn.execute("delete dividend")    
   conn.committrans
   conn.begintrans    
   myear  = request.form("myear")
   mrate = request.form("mrate")
   myr  = myear 
   nyr2 = (myr-1)&"0630"
   
  
   myr1 =dateserial((myr-1),6,30)
   myr2 = dateSerial(myr,6,1)
 
   dim sdate(12) 
   xx = 1
   yr = myr-1
   do while xx <=12
      mn = 6 + xx
      xyr = yr  
      if mn > 12 then
         mn = mn - 12
         xyr = yr + 1
      end if  
      sdate(xx)=dateSerial(xyr,mn,"01")
     
    xx=xx + 1
    loop


   gttlamt = 0
   set rs = server.createobject("ADODB.Recordset")
   sql ="select a.memno ,b.code,b.ldate,b.amount,a.mstatus from memmaster a , share b where  a.memno=b.memno  and  b.ldate<'"&myr2&"'  and (a.memdate< '"&myr2&"') and  ((a.wdate is null) or (a.wdate > '"&myr2&"' ))   order by a.memno,b.ldate,b.code,a.mstatus  "
   rs.open sql, conn, 1, 1
		
   dim  mPdamt(12) 
   xmemno= rs("memno")
   mstatus = rs("mstatus") 
   clsbal = 0
   mx = 0
   do while not rs.eof
     
      if xmemno <> rs("memno") then
         subttl = 0
         for i = 1 to 12
             subttl = subttl + round(mpdamt(i)*mrate/100/12  ,2)
     
         next 
        
         if subttl > 0 then
             if int(subttl*100)<>subttl*100 then
                if (subttl*100-int(subttl*100))>=0.5 then  
                    subttl = int(subttl*100)/100 + 0.01  
                 else
                    subttl = int(subttl*100)/100
                 end if 
             end if 
            gttlamt = gttlamt+ subttl 
            if mstatus="T" or mstatus="M" then
               conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'C' ,0) ")
            else
            if subttl<=100 then
                conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'S',0 ) ")  
            else     
                 conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'B',0 ) ")
            end if
            end if
         end if
         xmemno= rs("memno")
         mstatus = rs("mstatus") 
         clsbal = 0
         for i = 1 to 12
              mpdamt(i) = 0
         next 

      end if
      if rs("ldate") <=myr1 then
         select case left(rs("code"),1)
                case "0" ,"A","C"
                     if rs("code")<>"AI" then
                        clsbal = clsbal + rs("amount")
                     end if
               case  "B","G","H"
                       clsbal = clsbal - rs("amount")
         end select
         mpdamt(1) = clsbal
       else
        
          select case left(rs("code"),1)
                 case "0" ,"A","C"
                     if rs("code")<>"AI" then
                        if mx = 0 then
                            clsbal = clsbal + rs("amount")
                        else
                         
                                clsbal = clsbal + rs("amount")
                                mx = 0
                         
                        end if     
                     end if
                 case  "B"
                     clsbal = clsbal - rs("amount")
                                   mx = 1
                                   xmon = month(rs("ldate"))
                                   xyr  = year(rs("ldate"))
                                   xdd = mid("312831303130313130313031",(xmon-1)*2+1,2)
                                   if xmon =2 and int(xyr/4)=xyr/4 and int(xyr/100)=xyr/100 then
                                      xdd = 29
                                   end if
                                   chkdate = rs("ldate")
                                  if xmon < 7 then
                                     mon = xmon + 6
                                  else
                                     mon = xmon - 6
                                 end if  
                                 for i =  1 to mon
                                    mPdamt(i)= clsbal
                                 next 
                                         
                case  "G","H"
                        if mx = 0 then
                          clsbal = clsbal - rs("amount")     
                        else
                            if  rs("ldate")>=chkdate then
                                clsbal = clsbal - rs("amount")
                                mx = 0
                            end if 
                        end if     
         end select
         
         for i = 2 to 12
            if rs("ldate")< sdate(i) then
             mpdamt(i) = clsbal
            
            end if
         next         
       end if
 
      

   rs.movenext
   loop
   rs.close
 subttl = 0
   for i = 1 to 12
      subttl = subttl + mpdamt(i)*mrate/100/12  
     
   next 
        
         if subttl > 0 then
             if int(subttl*100)<>subttl*100 then
                if (subttk*100-int(subttl*100))>=0.5 then  
                    subttl = int(subttl*100)/100 + 0.01  
                 else
                    subttl = int(subttl*100)/100
                 end if 
             end if 
            gttlamt = gttlamt+subttl 
            if mstatus="T" or mstatus="M" then
               conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'C' ,0) ")
            else
            if subttl<=100 then
                conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'S',0 ) ")  
            else     
                 conn.execute("insert into dividend (memno,dividend,bank,deleted) values ( '"&xmemno&"',"&subttl&",'B',0 ) ")
            end if
            end if
         end if
   
    check = 1
  conn.committrans 
  
    msg="計算股息完成"    
else
       mrate =5.00
       myear =year(date()) 
       check=0
end if 

%>
<html>

<head>
<title>股息計算操作</title>

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

	if (formObj.myear.value==""){
		reqField=reqField+", 年份";
		if (!placeFocus)
			placeFocus=formObj.myear;
	}
	if (formObj.mrate.value==""){
		reqField=reqField+", 股息率";
		if (!placeFocus)
			placeFocus=formObj.mrate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.myear.focus()">
<!-- #include file="menu.asp" -->
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<br>
<center>
<h3>股息計算操作</h3>
<form name="form1" method="post" action="dvdcal.asp">
<input type="hidden" name="check value="<%=check%>">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">年份</td>
		<td width="10"></td>
		<td><input type="text" name="myear" value="<%=myear%>" size="4" maxlength="4" >
         </tr>
         <tr>
                <td class="b8">股息率</td>
		<td width="10"></td>
		<td><input type="text" name="mrate" value="<%=mrate%>" size="5" maxlength="5" >
		<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
                
	</tr>
<% if check=1 then %>
        <tr>
             <td class="b8">預計股息金額</td>   
             <td width="10"></td>
	     <td id ="ttlamt"><%=formatnumber(gttlamt,2)%></td> 
        </tr> 
<%end if%>
</table>
</form>
</center>
</body>
</html>
