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
		
                repaydate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2) &"/"&year(date())

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
                cashamt = 0
                intamt  = 0
                saveamt = 0    
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
               select case xstatus
                       case "呆帳"
                          mstatus="L"
                       case "冷戶"
                            mstatus="D"
                       case " IVA "
                            mstatus="V"
                       case "退社"
                          mstatus="C"
                       case "去世"
                            mstatus="B"   
                       case "破產"
                          mstatus="P"
                       case "正常"
                            mstatus="N"   
                      case "新戶"
                            mstatus="J"   
                      case "暫停銀行"
                          mstatus="H"
                       case "自動轉帳"
                            mstatus="A"
                       case "庫房,銀行"
                          mstatus="M"
                      case "庫房"
                            mstatus="T"   
                      case "問題貸款"
                            mstatus="F"   
                       case  "自動轉帳(股金)"                       
                            mstatus="0"   
                       case  "自動轉帳(股金,利息)"
                             mstatus="1"
                       case  "自動轉帳(股金,本金)"
                             mstatus="2"
                       case "自動轉帳(利息,本金)"
                             mstatus="3"
                 end select

                set rs = server.createobject("ADODB.Recordset")
		conn.begintrans

                 xrepaydate = right(repaydate,4)&"/"&mid(repaydate,4,2)&"/"&left(repaydate,2)

                               yy = cint(right(repaydate,4))
                               mm = cint(mid(repaydate,4,2))
                               dd = cint(left(repaydate,2))
                               select case mm
                                      case 1,3,5,7,8,10,12
                                           md = 31  
                                      case 4,6,9,11
                                           md =30
                                      case 2
                                          if int(yy/100)=(yy/100) and int(yy/4)=(yy/4)   then 
                                             md=29
                                          else
                                             md = 28
                                          end if
                               end select                     
                select case mstatus
                       case "V","N","H","F"
                           if cashamt <>"" then 
                              conn.execute("insert into loan (memno,lnnum,code,ldate,amount) values ('"&memno&"','"&lnnum&"','E3','"&xrepaydate&"',"&cashamt&") ")
			      conn.execute("update loanrec set bal= bal-"&cashamt&" where lnnum='"&lnnum&"' " )
                           end if

                           if intamt <> "" then
                              conn.execute("insert into loan (memno,lnnum,code,ldate,amount) values ('"&memno&"','"&lnnum&"','F3','"&xrepaydate&"',"&intamt&") ")
                           end if 
  
                       case "A","1","2","3","M","T","0"                                                         
   
                            if cashamt> 0 then

                                  xx = cashamt 
                                  CC = 0
                                  if xx >= pamt then
                                     set rs = server.createobject("ADODB.Recordset") 
                                     sql = "select * from loan where lnnum='"&lnnum&"'  and code='ME'   and pflag = 1 "
                                     rs.open sql, conn, 2, 2
                                     Do while not rs.eof  and cc = 0
                                        if xx - rs("bal") >= 0 then 
                                          conn.execute("update loan set pdate='"&xrepaydate&"'  where memno='"&memno&"' and code='ME' and pflag= 1  ")
                                          conn.execute("update loan set pflag=0 where memno='"&memno&"' and code='ME' and pflag= 1")
                                        else
                                          conn.execute("update loan set pdate='"&xrepaydate&"'  where memno='"&memno&"' and code='ME' and pflag= 1")
                                          conn.execute("update loan set bal= bal - "&xx&"   where memno='"&memno&"' and code='ME' and pflag= 1")
                                           cc = 0
                                        end if
                                        xx = xx - rs("bal")
                                      rs.movenext
                                      loop
                                      rs.close
                                  end if                        
                                  
    			            conn.execute("update loanrec set bal= bal-"&cashamt&" where lnnum='"&lnnum&"' " )
                                    conn.execute("update loanrec set cleardate='"&xrepaydate&"' where lnnum='"&lnnum&"' and bal=0 " )
                                    conn.execute("update loanrec set repaystat ='C' where lnnum='"&lnnum&"' and bal=0 " ) 
                                    if xx > 0 then
   			               conn.execute("insert into loan (memno,lnnum,code,ldate,amount,pflag,bal) values ('"&memno&"','"&lnnum&"','E3','"&xrepaydate&"',"&cashamt&",1,"&xx&" ) ")                           
                                    else
                                       conn.execute("insert into loan (memno,lnnum,code,ldate,amount,pflag,bal) values ('"&memno&"','"&lnnum&"','E3','"&xrepaydate&"',"&cashamt&",0,"&xx&" ) ")                           
                                    end if 
                               if intamt = 0  and cashamt > 0 then
                                   xintamt = round(cashamt*0.01*(dd/md),2)
                                   conn.execute("insert into loan (memno,lnnum,code,ldate,amount,pflag,bal ) values ('"&memno&"','"&lnnum&"','IF','"&xrepaydate&"',"&xintamt&",1,"&xintamt&" ) ")                                         
                                
                                end if                                       

                              end if
                          if intamt >0  then
                              xint = 1
                            
                                xx = intamt
                                if pint>0 then
                                     
                                     sql ="select * from loan where lnnum='"&lnnum&"'  and code='MF'   and pflag = 1 "
                                     rs.open sql, conn, 2, 2
                                     Do while not rs.eof  and cc = 0
                                        if xx - rs("bal") >= 0 then 
                                          conn.execute("update loan set pdate='"&xrepaydate&"'  where memno='"&memno&"' and code='Mf' and pflag= 1")
                                          conn.execute("update loan set pflag=0 where memno='"&memno&"' and code='MF' and pflag= 1")
                                        else
                                          conn.execute("update loan set pdate='"&xrepaydate&"'  where memno='"&memno&"' and code='Mf' and pflag= 1") 
                                          conn.execute("update loan set bal= bal - "&xx&"   where memno='"&memno&"' and code='ME' and pflag= 1")
                                           cc = 0
                                        end if
                                        xx = xx - rs("bal")
                                      rs.movenext
                                      loop
                                      rs.close 
                                end if                                                               
                                if xx > 0 then
                                   conn.execute("insert into loan (memno,lnnum,code,ldate,amount,pflag,bal ) values ('"&memno&"','"&lnnum&"','F3','"&xrepaydate&"',"&intamt&",1,"&xx&" ) ")                                      
                                else
                                    conn.execute("insert into loan (memno,lnnum,code,ldate,amount,pflag,bal ) values ('"&memno&"','"&lnnum&"','F3','"&xrepaydate&"',"&intamt&",0,"&xx&" ) ")                                      
                                end if
                              
                            else
                                 xint = 0                                  
                            end if
                      
                end select
                if saveamt>0 then
                                xx = saveamt
                                if xx>=samt then
                                    sql = "select * from share where memno='"&memno&"'  and code='AI'   and pflag = 1 "
                                    rs.open sql, conn, 2, 2
                                     Do while not rs.eof  and cc = 0
                                        if xx - rs("bal") >= 0 then 
                                          conn.execute("update share set pdate='"&xrepaydate&"'  where memno='"&memno&"' and code='AI' and pflag= 1")
                                          conn.execute("update share set pflag=0 where memno='"&memno&"' and code='AI' and pflag= 1")
                                        else
                                          conn.execute("update share set pdate='"&xrepaydate&"'  where memno='"&memno&"' and code='AI' and pflag= 1") 
                                          conn.execute("update share set bal= bal - "&xx&"   where memno='"&memno&"' and code='AI' and pflag= 1")
                                           cc = 0
                                        end if
                                        xx = xx - rs("bal")
                                      rs.movenext
                                      loop
                                      rs.close
                                     
                                  end if         
                                if xx > 0 then
                                   conn.execute("insert into share (memno,code,ldate,amount,pflag,bal) values ('"&memno&"','A3','"&xrepaydate&"',"&saveamt&",1,"&xx&"  ) ")    
                                else                                                                                                       
                                   conn.execute("insert into share (memno,code,ldate,amount,pflag,bal) values ('"&memno&"','A3','"&xrepaydate&"',"&saveamt&",0,"&xx&"  ) ")                                      
                                end if 
                              end if                                

                         





                            
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
<title>現金還款</title>
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
<form name="form1" method="post" action="repayloan.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<div><center><font size="3">現金還款</font></center></div>
<center>
<table border="0" cellspacing="0" cellpadding="0">
       <tr>
		<td width="500" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">社員號碼</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>>
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
				<td><input type="text" name="memName" value="<%=memName%>" size="30"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>
                       </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left"></td>
				<td width=50></td>
				<td><input type="text" name="memcName" value="<%=memcName%>" size="30"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員現狀</td>
		<td width=45></td>
		<td id="xstatus" ><%=xstatus%></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">還款日期</td>
		<td width=50></td>
		<td><input type="text" name="repaydate" value="<%=repaydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
	</tr>


	<tr>
               <td width=30></td>
		<td class="b12" align="left">還款金額</td>
		<td width=50></td>
		<td><input type="text" name="cashamt" value="<%=cashamt%>" size="10" ></td>
	</tr>

       
	<tr>
               <td width=30></td>
		<td class="b12" align="left">還款利息</td>
		<td width=50></td>
		<td><input type="text" name="intamt" value="<%=intamt%>" size="10"  ></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">還款股金</td>
		<td width=50></td>
		<td><input type="text" name="saveamt" value="<%=saveamt%>" size="10"  ></td>
	</tr>

	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if id <> "" then %>
               
		<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
		<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
		<input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
		<%end if%>			
		<% end if %>
               
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
