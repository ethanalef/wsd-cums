<!-- #include file="../conn.asp" -->
<!-- #include file="cutpro.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%

if request.form("bye") <> "" then
   response.redirect "main.asp"
end if

if request.form("clrScr") <> "" then
     memno=""
     amount=""
     memName =""
     memcName =""
     id =""

end if

if request.form("Search") <> "" then 

	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

   if memno <>"" then          
        set rs = conn.execute("select memno,memName,memcName,mstatus from memmaster where memno='"&memno&"'  ")
			For Each Field in rs.fields			
			TheString = Field.name & "= rs(""" & Field.name & """)"
			Execute(TheString)
			Next

        if not rs.eof then 
           work = 1
           id = memno 
	     select case mstatus
	          case "C"	
               id=""
               msg = "** 社員巳退社 **"		
          case "P"
                id = "" 
		msg = "** 社員巳去世 **"
          case "B"
                id =""
		msg = "** 社員巳破產 **"
          case "L"
                id=""
		msg = "** 社員在呆帳中 **"
          case "V"
                id =""
		msg = "** 社員在 ＩＶＡ 中 **"
          case "F"
                id ="" 
		msg = "** 社員還款有問題 **"
   end select
      
        else
          msg = "社員不存在"
       end if
    rs.close
         
   end if 
   unpttl = 0
   if msg="" then
   
      set rs= conn.execute("select * from share where memno='"&memno&"' and pflag='1'  ")
      do while not rs.eof
         select case rs("code")
                case "AI"  
                   unpttl = unpttl + rs("amount")      
                case  else
         end select
         rs.movenext
      loop
      rs.close
  end if   
  if mstatus="D" then
      set rs= conn.execute("select * from share where memno='"&memno&"'   ")
      do while not rs.eof
         select case left(rs("code"),1)
                case "A" ,"C","0" 
                    if rs("code") <>"AI"  then
                   
                       bal = bal + rs("amount")   
                     end if    
                case  "B","G","G","M"
                      bal = bal - rs("amount")  
                     
         end select
         rs.movenext
      loop
      rs.close     
      if bal >= 300 then 
         amount = 300 
       else
          amount= bal
       end if
    end if
else
id = ""

end if

if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
         sxdate=right(ldate,4)&"/"&mid(ldate,4,2)&"/"&left(ldate,2)
               select case item
                       case "現金"
                             addUserLog "增加 現金帳  日期 : "&ldate&"(金額) $ "&amount
                             acode ="A3"
                       case "調整"
                              addUserLog "增加 調整帳  日期 : "&ldate&"(金額) $ "&amount
                             acode ="A7"
                       case "保險金"
                              addUserLog "增加 保險金帳  日期 : "&ldate&"(金額) $ "&amount
                             acode ="A4"
                       case "冷戶費"
                              addUserLog "增加 保險金帳  日期 : "&ldate&"(金額) $ "&amount
                             acode ="MF"
                             
                  end select
        conn.begintrans   

        if unpttl > 0 and  (amount-unpttl)>=0 then
           xamt = amount
           set rs = server.createobject("ADODB.Recordset")
           sql = "select * from share where memno='"&memno&"' and pflag = 1 "
           rs.open sql, conn, 1, 1
           if not rs.eof then 
              do while  not rs.eof 
                 if xamt >= rs("bal") then
                                     
                    conn.execute("update share set pflag = 0 where memno='"&memno&"' and code='AI' and ldate='"&rs("ldate")&"' and pflag=1 ")
                     
                else
                   conn.execute("update share set bal = bal - '"&xamt&"' where memno = '"&memno&"' and code='AI' and ldate= '"&rs("ldate")&"' and pflag = 1 ")
         
               end if
           
              xamt  = xamt - rs("bal")
              if xamt = 0 then
                 rs.movelast
              end if
          rs.movenext  
          loop  
          end if
          rs.close                     
       end if
     
    
         conn.execute("insert into share (memno,ldate,code,amount) values ('"&memno&"','"&sxdate&"','"&acode&"',"&amount&") ")
  
          
        
        conn.committrans
 
	msg = "紀錄已更新"
        id = ""
else
   msg=""
end if

 
     ldate=RIGHT("0"&day(date()),2)&"/"&RIGHT("0"&month(date()),2)&"/"&year(date())          
     chkdate=RIGHT("0"&day(date()),2)&"/"&RIGHT("0"&month(date()),2)&"/"&(year(date())-1)



	

%>
<html>
<head>
<title>現金存款建立</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}
function valDate(M, D, Y){
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31) || (M < 1) || (M > 12) || (Y < 0))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
  return true;
};

function formatDate(dateform){


	    formObj=document.form1;    
            sMn = parseInt(formObj.lastmonth.value)
            sYr = parseInt(formObj.lastyear.value)
            spass   = parseInt(formObj.spass.value)

  cDate = dateform.value;
  dSize = cDate.length;
  if (dSize!=0){
    sCount= 0;
    for(var i=0; i < dSize; i++)
      (cDate.substr(i,1) == "/") ? sCount++ : sCount;
    if (sCount == 2){
		ySize = cDate.substring(cDate.lastIndexOf("/")+1,dSize).length;
		if (ySize<2 || ySize>4 || ySize == 3){
		  return false;
		 }
		idxBarI = cDate.indexOf("/");
		idxBarII = cDate.lastIndexOf("/");
		strD = cDate.substring(0,idxBarI);
		strM = cDate.substring(idxBarI+1,idxBarII);
		strY = cDate.substring(idxBarII+1,dSize);
		strM = (strM.length < 2 ? '0'+strM : strM);
		strD = (strD.length < 2 ? '0'+strD : strD);
		if(strY.length == 2)
		  strY = (strY > 50  ? '19'+strY : '20'+strY);
    }else{
    	if (dSize != 8)
			return false;
		strD = cDate.substring(0,2);
		strM = cDate.substring(2,4);
		strY = cDate.substring(4,8);
    }
    dateform.value = strD+'/'+strM+'/'+strY;
    if (!valDate(strM, strD, strY))
      return false;
    else

     Mn = parseInt(strM)
      Yr = parseInt(strY)
      if (((Mn<=sMn)&&(Yr==sYr))||(Yr<sYr)){
         return false ;
      }else{      
         return true;
      }
  }
}



function memberChange(){
	if (document.form1.memNo.value==''){
		document.form1.memName.value=''
		document.all.tags( "td" )['memName'].innerHTML=''
		
		

        
        }
}


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;



 	if (formObj.amount.value==""){
		reqField=reqField+", 存款金額";
		if (!placeFocus)
			placeFocus=formObj.amount;
	}

	if (!formatDate(formObj.ldate)){
		reqField=reqField+", 存款日期";
		if (!placeFocus)
			placeFocus=formObj.ldate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNo.focus()">
<!-- #include file="menu.asp" -->
<div><center><font size="3">現金存款建立</font></center></div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="savecash.asp">
<table border="0" cellspacing="0" cellpadding="0">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memcName" value="<%=memcName%>">
<input type="hidden" name="monthsave" value="<%=monthsave%>">
<input type="hidden" name="monthssave" value="<%=monthssave%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="unpttl" value="<%=unpttl%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<table border="0" cellspacing="0" cellpadding="0">

				<tr>
					<td class="b8" align="right">社員編號</td>
					<td width=10></td>
					<td>
						<input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10" onchange="memberChange()"<%if id<>"" then response.write " onfocus=""form1.ldate.focus();""" end if%>>
                                                <%if id ="" then %>
						<input type="button" value="選擇" onclick="popup('pop_srhMemnoM.asp')" class="sbttn">
                                                <input type="submit" value="搜尋" name="Search" class ="Sbttn"> 
                                                <%end if%>
					</td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">姓名</td>
					<td width=10></td>
					<td id="memname"><%=memname%></td>
                                
				</tr>
				<tr height="22">
					<td class="b8" align="right"></td>
					<td width=10></td>
					
                                        <td id="memcname"><%=memcname%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">脫期股金</td>
					<td width=10></td>
					
                                        <td id="unpttl"><%=formatnumber(unpttl,2)%></td>
				</tr>				
				<tr height="22">
					<td class="b8" align="right">日期</td>
					<td width=10></td>
					<td><input type="text" name="ldate"  value="<%=ldate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};">(dd/mm/yyyy)</td>
				</tr>
                                <tr height="22">
                                <td class="b8" align="right">類別</td>
                               
                                <td width="10"></td>
    		                <td>
			        <select name="item">
			        <option<%if item="C" then response.write " selected" end if%>>現金</option>
			        <option<%if item="A" then response.write " selected" end if%>>調整</option>
                                <option<%if item="I" then response.write " selected" end if%>>保險金</option>
                                <option<%if item="I" then response.write " selected" end if%>>冷戶費</option>
			        </select>
		                </td>   
                                </tr>  
                                <tr>
				<td class="b8" align="right">金額</td>
				<td width=10></td>
				<td><input type="text" name="amount" value="<%=amount%>" size="10" maxlength="10" </td>
                                </tr>
          
		<tr>
					<td colspan="3" align="right">
					<%if id<>"" then %>
						<%if session("userLevel")<>2 and session("userLevel")<>1 and session("userLevel")<>4 then%>
						<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
						<%end if%>
						
						<input type="button" value="查詢個人帳" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value  )" class="sbttn">
                                                <%end if%>
						<input type="submit" value="取消" name="clrSrc" class="sbttn">
						<input type="submit" value="返回" name="bye" class="sbttn">
				</td>
				</tr>

</table>
<br>

</center>
</form>
</body>
</html>
