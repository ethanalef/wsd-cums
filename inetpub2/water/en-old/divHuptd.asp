<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
end if
if request.form("cancel") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

        conn.begintrans
        conn.execute("delete  share where memno='"&memno&"' and code='CH'  ")        
 
         conn.committrans
       
             
		
		msg = "紀錄已刪除"

        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)        
end if

if request.form("Search")<>"" or id <>""  then
 cashamt=0
                      intamt=0
                      saveamt=0
                     
        msg=""
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

                set rs=conn.execute("select a.memno,a.memname,a.memcname,a.mstatus,b.* from memmaster a,share b where a.memno='"&memno&"'  and b.memno=a.memno and b.code='H' ")
                if not rs.eof then
                   For Each Field in rs.fields 
		   TheString = Field.name & "= rs(""" & Field.name & """)"
	           Execute(TheString)
		   Next
                   id = memno
 
                   ldate=rs("ldate")
                      
                   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
                   paydate = right("0"&day(ldate),2)&"/"&right("0"&month(ldate),2)&"/"&(year(ldate))    
                   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)                         
                else
           
                   id = "" 
                   msg =" 不是暫停派息社員 "
                   memno = ""                           
                end if
                rs.close
    
             

 
                opt = 0

else 




if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

        yy  = right(paydate,4)
        mm  = mid(paydate,4,2)
        dd  = left(paydate,2)  
        
        xxdate= dateSerial(yy,mm,dd)
         conn.begintrans
         select case paid 
                case "S"
                 conn.execute("insert into share (memno , ldate , code , amount ) values ( '"&memno&"' ,'"&xxdate&"' ,'C0' , "&amount&" ) ")  
                case "B"
                 conn.execute("insert into share (memno , ldate , code , amount ) values ( '"&memno&"' ,'"&xxdate&"' ,'C0' , "&amount&" ) ")        
                  conn.execute("insert into share (memno , ldate , code , amount ) values ( '"&memno&"' ,'"&xxdate&"' ,'C1' , "&amount&"*-1 ) ")
                case "C"
                  conn.execute("insert into share (memno , ldate , code , amount ) values ( '"&memno&"' ,'"&xxdate&"' ,'C0' , "&amount&" ) ")        
                  conn.execute("insert into share (memno , ldate , code , amount ) values ( '"&memno&"' ,'"&xxdate&"' ,'C3' , "&amount&"*-1 ) ")
        end select 
        conn.execute("delete  share where memno='"&memno&"' and code='CH'  ")  
        conn.committrans
       

 
                     
                                             
		
		msg = "紀錄已更新"

        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
       
else
   id = ""
   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
   
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
end if
end if
%>
<html>
<head>
<title>暫停派息過數</title>
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
      return true;
  }
}






function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.memNo.value==""){
		reqField=reqField+", 社員編號";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}


	if (formObj.paydate.value==""){
		reqField=reqField+", 派息日期";
		if (!placeFocus)
			placeFocus=formObj.paydate;
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
<DIV>




<!-- #include file="menu.asp" --> 
<br>
<form name="form1" method="post" action="divHuptd.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xbank" value="<%=xbank%>">


<div><center><font size="3">暫停派息過數</font></center></div>
<br>

<center>
<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<BR>
<% end if%>

<table border="0" cellspacing="0" cellpadding="0">
 
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">社員號碼</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.paydate.focus();""" end if%>>
			<%if id = "" then %>	
		         <input type="submit"  value="搜尋" name="Search" class ="Sbttn">
			<% end if %>
                        </TD>
                        </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left">社員名稱</td>
				<td width=50></td>
				<td><input type="text" name="memName" value="<%=memName%>" size="30"<%if id<>"" then response.write " onfocus=""form1.paydate.focus();""" end if%>></td> 
			</tr>
                       </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left"></td>
				<td width=50></td>
				<td><input type="text" name="memcName" value="<%=memcName%>" size="30"<%if id<>"" then response.write " onfocus=""form1.paydate.focus();""" end if%>></td> 
			</tr>

			<tr>
                		<td width=30></td>
				<td class="b12" align="left">股息金額</td>
				<td width=50></td>
				<td><input type="text" name="amount" value="<%=amount%>" size="30"<%if id<>"" then response.write " onfocus=""form1.paydate.focus();""" end if%>></td> 
			</tr>

			<tr>
                		<td width=30></td>
				<td class="b12" align="left">股息金額</td>
				<td width=50></td>
                                <td> 
                                     <INPUT TYPE=RADIO NAME="Paid" VALUE="S">Share<BR>
                                     <INPUT TYPE=RADIO NAME="Paid" VALUE="C">Cheque/Cash<BR>
                                      <INPUT TYPE=RADIO NAME="Paid" VALUE="B">Bank
                                </td>
                      	</tr>
			
	<tr>
               <td width=30></td>
		<td class="b12" align="left">派息日期</td>
		<td width=50></td>
		<td><input type="text" name="paydate" value="<%=paydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
                    
                </td>                 
 	</tr>
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if id <> "" then %>
               
		<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
		<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
		<input type="submit" value="刪除" onclick="confirm('確定刪除?')" name="cancel" class="sbttn">
		<input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
               
		<%end if%>			
		<% end if %>
               
		<input type="submit" value="取消" name="bye" class="sbttn">
		<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
 
</table>       
</center>
</form>
</body>
</html>
