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
    pint = 0
   pamt  = 0 
   intamt = 0
   cashamt = 0
   lostamt = 0
   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
end if


if request.form("Search")<>"" or id <>""  then
        msg=""
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        if id<>"" then 
		sql = "select * from loanrec where lnnum='"& xlnnum & "' "
        else
                set rs=conn.execute("select memno,memname,memcname,mstatus from memmaster where memno='"&memno&"' ")
                if not rs.eof then
                   For Each Field in rs.fields 
		   TheString = Field.name & "= rs(""" & Field.name & """)"
	           Execute(TheString)
		   Next
                id = memno  

                else
                    
                   msg ="貸款號碼不存在 "
                end if 
                rs.close
                 sql ="select * from loanrec where repaystat='N' and memno="&memno
        end if
                select case mstatus
                       case "L"
                           xstatus= "呆帳"
                       case  "D"
                           xstatus="冷戶"
                       
                       case  "V"
                           xstatus= " IVA "
                         
                       case  "C"
                             xstatus= "退社"
             
                       case  "B"
                             xstatus= "去世"
                         
                       case  "P"
                            xstatus="破產"
                    
                       case  "N"
                            xstatus= "正常"
                        
                      case  "J"
                            xstatus= "新戶"
                       
                      case "H"
                          xstatus= "暫停銀行"
                      
                       case  "A"
                            xstatus="自動轉帳"

                       case  "0"
                            xstatus="自動轉帳(股金)"                       
                       case  "1"
                            xstatus="自動轉帳(股金,利息)"
                       case  "Z"
                            xstatus="自動轉帳(股金,本金)"
                       case "3"
                             xstatus="自動轉帳(利息,本金)"
                       case  "M"
                           xstatus = "庫房,銀行"
                      
                      case  "T"
                            xstatus= "庫房"
                     case "F"
                          xstatus =  "問題貸款"
                end select
        if msg="" then
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if not rs.eof then

			For Each Field in rs.fields
			if Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		
                delydate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2) &"/"&year(date())
                if delyflag  <>"Y" then
                   delyflag = "N"
                   months = "3" 
                   delydate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2) &"/"&year(date())
                else
                    delydate=right("0"&day(delydate),2)&"/"&right("0"&month(delydate),2) &"/"&year(delydate) 
                end if
  
                id = memno
                rs.close 
		pint = 0
		pamt  = 0 
                yy = year(date())
                mm  =month(date())-1
		sql = "select *  from loan where memno='"& memno & "'  and code='ME' and pflag= 1 "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while  not rs.eof 
                      pamt = pamt + rs("bal")
                rs.movenext
                loop             
                rs.close
		sql = "select * from loan where memno ='"& memno & "'  and code= 'MF' and pflag = 1   "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                      pint = pint + rs("bal")       
                rs.movenext
                loop
                rs.close
                samt = 0 
		sql = "select * from share where memno ='"& memno & "'  and code= 'AI' and pflag = 1   "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                      samt = samt + rs("bal")       
                rs.movenext
                loop
                rs.close
                ttlamt = bal+pint+pamt+samt                        
   
                end if      
                else
                  msg = "貸款號碼不存在 "

                end if   	
end if



if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
 

		conn.begintrans

                 xdelydate = right(delydate,4)&"/"&mid(delydate,4,2)&"/"&left(delydate,2)
                pos = instr(months,",")
                mons = right(months,3)
 
                               select case mons 
                                      case "一個月"
                                           conn.execute("update loanrec set months = '1' where lnnum = '"&lnnum&"' ")
                                           conn.execute("update loanrec set chkmon= 1 where lnnum = '"&lnnum&"'  ")
                                      case "二個月"
                                           conn.execute("update loanrec set months = '2' where lnnum = '"&lnnum&"' ")
                                           conn.execute("update loanrec set chkmon= 2 where lnnum = '"&lnnum&"'  ")
                                      case "三個月"
                                          conn.execute("update loanrec set months = '3' where lnnum = '"&lnnum&"' ")
                                          conn.execute("update loanrec set chkmon= 3 where lnnum = '"&lnnum&"'  ")
                               end select                     
             
               
                conn.execute("update loanrec set delydate = '"&xdelydate&"' where lnnum = '"&lnnum&"' ")
                
                conn.execute("update loanrec set delyflag =  'Y' where lnnum = '"&lnnum&"'  ")
		conn.committrans
		msg = "紀錄已更新"

        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

 
       
else
   mstatus=""
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)          
   
end if
%>
<html>
<head>
<title>延期操作</title>
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



	if (!formatDate(formObj.repaydate)){
		reqField=reqField+", 還款日期";
		if (!placeFocus)
			placeFocus=formObj.repaydate;
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


<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<% end if%>

<br>
<form name="form1" method="post" action="delayPro.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="months" value="<%=months%>">
<div><center><font size="3">延期操作</font></center></div>
<center>
<table border="0" cellspacing="0" cellpadding="0">
       <tr>
		<td width="500" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">社員號碼</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.delydate.focus();""" end if%>>
			<%if id = "" then %>
			<input type="button" value="選擇"  onclick="popup('pop_srhLoan.asp?key='+document.form1.memNo.value)" class="sbttn"  >
			<input type="submit" value="搜尋" name="Search" class ="Sbttn">
			<% end if %>
                        </TD>
                        </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left">社員名稱</td>
				<td width=50></td>
				<td><input type="text" name="memName" value="<%=memName%>" size="30"<%if xlnnum<>"" then response.write " onfocus=""form1.delydate.focus();""" end if%>></td> 
			</tr>
                       </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left"></td>
				<td width=50></td>
				<td><input type="text" name="memcName" value="<%=memcName%>" size="30"<%if xlnnum<>"" then response.write " onfocus=""form1.delydate.focus();""" end if%>></td> 
			</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員現狀</td>
		<td width=45></td>
		<td id="xstatus" ><%=xstatus%></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">開始延期日期</td>
		<td width=50></td>
		<td><input type="text" name="delydate" value="<%=delydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
	</tr>


	<tr>
               <td width=30></td>
		<td class="b12" align="left">延期期數</td>
		<td width=50></td>
		<td>
			<select name="months">                       
			<option<%if months = "1" then response.write " selected" end if%>> 一個月</option>
                        <option<%if months = "2" then response.write " selected" end if%>> 二個月</option>
                        <option<%if months = "3" then response.write " selected" end if%>> 三個月</option>
			</select>
		</td>


	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<%if  id  <>""  then%>
		<%if  delyflag <> "Y"   then%>
		<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
                <% end if %>
		<input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">											               
                <%end if %>   
		<input type="submit" value="取消" name="bye" class="sbttn">
		<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
      </table>
      </td>
	<td width="400" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
                		<td width=30></td>
				<td class="b12" align="left">貸款號碼</td>
				<td width=50></td>
				<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">取票日期</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款結餘</td>
		<td width=50></td>
		<td><input type="text" name="bal" value="<%=bal%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">脫期本金</td>
		<td width=50></td>
		<td><input type="text" name="pamt" value="<%=pamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">脫期股金</td>
		<td width=50></td>
		<td><input type="text" name="samt" value="<%=samt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">脫期利息</td>
		<td width=50></td>
		<td><input type="text" name="pint" value="<%=pint%>" size="10"></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">總欠款</td>
		<td width=50></td>
		<td><input type="text" name="ttlamt" value="<%=ttlamt%>" size="10"></td> 
	</tr>
        </table>
        </td>  
   </tr>
</table>       
</center>
</form>
</body>
</html>
