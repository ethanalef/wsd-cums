<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if
opt = 0



  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next
   pint = 0
   pamt  = 0 

    memno = ""
   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
else
 memno      = request.form("memno") 
end if

pass = 0
opt = 0
if request.form("Search")<>"" or id <>""  then

                     
                     
        msg=""
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

                set rs=conn.execute("select memno,memname,memcname,mstatus from memmaster where memno='"&memno&"' and mstatus ='F'  ")
                if not rs.eof then
                   For Each Field in rs.fields 
		   TheString = Field.name & "= rs(""" & Field.name & """)"
	           Execute(TheString)
		   Next
                   id = memno  
                   
                else
           
                   id = "" 
                   msg =" 不是問顯社員 "
                   memno = ""                           
                end if
                rs.close
       if msg="" then
                
		set rs = server.createobject("ADODB.Recordset")
                sql = "select * from memmaster  where memno='"&memno&"'  "
		rs.open sql, conn,1,1
		if not rs.eof then
                   select case rs("mstatus")
                          case "A","0","1","2"
                               samt = rs("monthsave")
                          case "T","M"
                               samt = rs("monthssave")
                  end select
                end if
                rs.close 
      
              
		set rs = server.createobject("ADODB.Recordset")
                sql = "select * from loanrec where memno='"&memno&"' and repaystat='N' "
		rs.open sql, conn,1,1
		if not rs.eof then

			For Each Field in rs.fields
			if Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		xlnnum = lnnum
                repaydate=""
                
                id = memno
                rs.close
		pint = 0
		pamt  = 0 
                yy = year(date())
                mm  =month(date())-1
                pamt = 0
		sql1 = "select *  from loan where memno='"& memno & "'  and code='DE' and pflag= 1 "
		set rs1 = server.createobject("ADODB.Recordset")
		rs1.open sql1, conn,2,2
                do while  not rs1.eof 
                      pamt = pamt + rs1("bal")
                rs1.movenext
                loop             
                rs1.close
                pint = 0
		sql1 = "select * from loan where memno ='"& memno & "'  and code= 'DF' and pflag = 1   "
		set rs1 = server.createobject("ADODB.Recordset")
		rs1.open sql1, conn
                do while not  rs1.eof 
                      pint = pint + rs1("bal")       
                rs1.movenext
                loop
                rs1.close
                samt = 0 
		sql1 = "select * from share where memno ='"& memno & "'  and code= 'AI' and pflag = 1   "
		set rs1 = server.createobject("ADODB.Recordset")
		rs1.open sql1, conn
                do while not  rs1.eof 
                      samt = samt + rs1("bal")       
                rs1.movenext
                loop
                rs1.close
                ttlamt = bal+pint+pamt+samt          
                end if      
                else
                  msg = "貸款號碼不存在 "

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

                       case  "S"
                            xstatus="自動轉帳(股金)"                       
                       case  "I"
                            xstatus="自動轉帳(股金,利息)"
                       case  "Z"
                            xstatus="自動轉帳(股金,本金)"
                       case  "T"
                           xstatus = "庫房,銀行"
                      
                      case  "M"
                            xstatus= "庫房"
                         
                end select
                opt = 0


		sql1 = "select * from autopay where memno ='"& memno & "'  and status ='F' "
		set rs1 = server.createobject("ADODB.Recordset")
		rs1.open sql1, conn
                if not rs1.eof then
                    opt = 1
                   do while not  rs1.eof 
                       select case rs1("code")
                              case "A1"
                                   saveamt = rs1("bankin")                                  
                              case "A2"
                                   msaveamt = rs1("bankin")                                  

                              case "E1"
                                   cashamt = rs1("bankin")   
                                    mstatus=rs1("mstatus")                             
                              case "E2"
                                   
                                   mcashamt = rs1("bankin")   
                                   mstatus=rs1("mstatus")                                
                              case "F2"
                                   mintamt = rs1("bankin")
                                   
                              case "F1"
                                   intamt = rs1("bankin")   
                       end select
                     upflag=rs1("upflag") 
                     repaydate=dmy(rs1("adate"))
                    rs1.movenext
                    loop
                    rs1.close
                    target = upflag
                    xtarget = upflag
                    pass = 1    
      end if
               
         
else




if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next


   select case target
          case "銀行"
                target="B"
                
          case "庫房"
                target="T"
          case "庫房/銀行"
                target="M"
   end select
    xtarget = target 
 
        	conn.begintrans
                
                pdate = year(date())&"/"&month(date())&"/"&day(date())
                mdate = right(repaydate,4)&"/"&mid(repaydate,4,2)&"/"&left(repaydate,2) 
                xint = 0
                reamt = 0 
                If bal<>"" then
                xint = bal *.01
               
                reamt = cint(monthrepay)     

               END IF


                if trim(mintamt) <> "" then
                            
                              if opt = 0 then 
                               conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,pdate,status,flag,deleted,pflag ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F2',"&mintamt&","&xint&",'"&pdate&"','F' ,'N',0,0 ) ")                     
                               else
                                   conn.execute("update autopay set bankin = "&mintamt&" where memno='"&memno&"' and status='F' and code='F2' ") 
                               end if
                end if 

                if trim(intamt) <> "" then 
                  
                                                           
                               if opt = 0 then
                                   conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,pdate,status,flag,deleted ,pflag ) values ('"&memno&"','"&mdate&"','"&lnnum&"','F1',"&intamt&","&intamt&",'"&pdate&"','F','N',0,0 ) ")   
                               else
                                   conn.execute("update autopay set bankin = '"&intamt&"' where memno='"&memno&"' and status='F' and code='F1' ") 
                               end if    
                         
                  
                end if
                if trim( mcashamt) <> "" then
                               if opt = 0 then
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,pdate,status,flag,deleted,pflag ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E2',"&mcashamt&","&mcashamt&",'"&pdate&"','F','N',0,0 ) ") 
                               else
                                   conn.execute("update autopay set bankin = "&mcashamt&" where memno='"&memno&"' and status='F' and code='E2' ") 
                               end if  
                end if
                if trim(cashamt) <> "" then
                  
                      
                                if opt = 0 then 
                                conn.execute("insert into autopay (memno,adate,lnnum,code,bankin,curamt,pdate,status,flag,deleted,pflag  ) values ('"&memno&"','"&mdate&"','"&lnnum&"','E1',"&cashamt&","&cashamt&",'"&pdate&"','F','N',0,0 ) ") 
                               else
                                   conn.execute("update autopay set bankin = "&cashamt&" where memno='"&memno&"' and status='F' and code='E1' ") 
                               end if
                         
                end if            
                if trim(saveamt)  <> "" then  
                  
                        
                              if opt = 0 then
                                   conn.execute("insert into autopay (memno,adate,code,bankin,curamt,pdate,status,flag,deleted,pflag  ) values ('"&memno&"','"&mdate&"','A1',"&saveamt&","&saveamt&",'"&pdate&"','F','N',0 , 0 ) ") 
                               else
                                   conn.execute("update autopay set bankin = "&saveamt&" where memno='"&memno&"' and status='F' and code='A1' ") 
                               end if  
 
                  
                end if              
                if trim( msaveamt) <> "" then

                         
                               if opt = 0 then 
                               conn.execute("insert into autopay (memno,adate,code,bankin,curamt,pdate,status,flag,deleted,pflag  ) values ('"&memno&"','"&mdate&"','A2',"&msaveamt&","&msaveamt&",'"&pdate&"','F','N',0 ,0 ) ") 
                               else
                                   conn.execute("update autopay set bankin = "&msaveamt&" where memno='"&memno&"' and status='F' and code='A2' ") 
                               end if

               end if
                if instr(mstatus,"呆帳")> 0 then 
                             conn.execute("update autopay  set mstatus='L' where memno='"&memno&"' ")
                end if
                if instr(mstatus,"冷戶")> 0 then  
			     conn.execute("update autopay set mstatus='D' where memno='"&memno&"' ")
                end if
                if instr(mstatus,"IVA")> 0 then 
                           conn.execute("update autopay set mstatus='V' where memno='"&memno&"' ")
                 end if
                 if instr(mstatus,"退社" )> 0 then 
			   conn.execute("update autopay set mstatus='C' where memno='"&memno&"' ")
                end if
                if instr(mstatus, "去世" )> 0 then 
			  conn.execute("update autopay set mstatus='P' where memno='"&memno&"' ")
                end if 
                if instr(mstatus,"破產")> 0 then 
			  conn.execute("update autopay set mstatus='B' where memno='"&memno&"' ")
                end if 
                 if instr(mstatus,"正常")> 0 then 
			    conn.execute("update autopay set mstatus='N' where memno='"&memno&"' ")
                 end if
                if instr(mstatus, "新戶")> 0 then 
                           conn.execute("update autopay set mstatus='J' where memno='"&memno&"' ")
                end if
                if instr(mstatus, "庫房" )> 0 then 
                          conn.execute("update autopay set mstatus='T' where memno='"&memno&"' ")
                 end if
                 if instr(mstatus,"暫停銀行")> 0 then 
			   conn.execute("update memmaster set mstatus='H' where memno='"&memno&"' ")
                 end if
                 if instr(mstatus,"自動轉帳(ALL)")> 0 then 
			   conn.execute("update autopay set mstatus='A' where memno='"&memno&"' ")
                  end if
                  if instr(mstatus,"自動轉帳(股金)")> 0 then 
			    conn.execute("update autopay set mstatus='0' where memno='"&memno&"' ")
                  end if
                 if instr(mstatus,"自動轉帳(股金,利息)")> 0 then 
			   conn.execute("update autopay set mstatus='1' where memno='"&memno&"' ")
                  end if
                  if instr(mstatus, "自動轉帳(股金,本金)" )> 0 then                         
			    conn.execute("update autopay set mstatus='2' where memno='"&memno&"' ")
                  end if
                  if instr(mstatus,"自動轉帳(利息,本金)")> 0 then                          
			    conn.execute("update autopay set mstatus='3' where memno='"&memno&"' ")
                  end if
                  if instr(mstatus,"庫房,銀行")> 0 then 
			     conn.execute("update autopay set mstatus='M' where memno='"&memno&"' ")
                  end if
                  if instr(mstatus,"特別個案")> 0 then   
			    conn.execute("update autopay set mstatus='F' where memno='"&memno&"' ")
                  end if 
                  if instr(mstatus,"終止社籍轉帳")> 0 then 
                            conn.execute("update autopay set mstatus='8' where memno='"&memno&"' ")
                  end if 
                  if instr(mstatus, "終止社籍正常")> 0 then 
                            conn.execute("update autopay set mstatus='9' where memno='"&memno&"' ")                          
                  end if 
   select case target
          case "銀行"
                target="B"
                
          case "庫房"
                target="T"
          case "庫房/銀行"
                target="M"
   end select
    
              if opt = 0 then
                conn.execute("update autopay  set upflag='"&target&"'   where memno='"&memno&"' ")
            end if
		conn.committrans
		msg = "紀錄已更新"

        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
        opt = 0
  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 repaydate=""
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
       
else
                if request.form("delrec") <>"" then
                    memno = request.form("memno")
                    opt = 0
                        conn.execute("delete autopay where memno= '"&memno&"' ")
                    memno =""  
else


if memno<>""  and pass = 0  then
 
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

   id = memno
   select case target
          case "銀行"
                target="B"
                
          case "庫房"
                target="T"
          case "庫房/銀行"
                target="M"
   end select
    xtarget = target 

    
end if

   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
   repaydate= ""
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
 
end if
end if
end if
%>
<html>
<head>
<title>特別個案輸入操作</title>
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
<br>
<form name="form1" method="post" action="Mautopay.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xtarget" value="<%=xtarget%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="opt" value="<%=opt%>">
<input type="hidden" name="monthrepay" value="<%=monthrepay%>">
<div><center><font size="3">特別個案轉帳輸入操作</font></center></div>
<center>
<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<% end if%>
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
		         <input type="submit"  value="搜尋" name="Search" class ="Sbttn">
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
               <td class="b12" align="left">轉帳單位</td>
                <td width=50></td>

               <td> 
<% if opt = 0 then %>
			<select name="Target" onchange="this.form.submit()"    >

			<option<%if Target="B" then response.write " selected" end if%>>銀行</option>
			<option<%if Target="T" then response.write " selected" end if%>>庫房</option>
                        <option<%if Target="M" then response.write " selected" end if%>>庫房/銀行</option>

			</select>
<%else
           select case target 
                  case  "B"
                     response.write("銀行")
                  case "T"
                      response.write("庫房")
                  case "M"
                       response.write("庫房/銀行")
           end select 
end if
 
%>               </td> 







       
	</tr>
	<tr>
               <td width=50></td>
		<td class="b12" align="left">轉帳日期</td>
		<td width=50></td>
		<td><input type="text" name="repaydate" value="<%=repaydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
	</tr>

<% if (xtarget="M" or xtarget="T") and lnnum<>"" then %>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">庫房還款利息</td>
		<td width=50></td>
		<td><input type="text" name="mintamt" value="<%=mintamt%>" size="10"  ></td>
	</tr>
	<tr>
               <td width=50></td>
		<td class="b12" align="left">庫房還款本金</td>
		<td width=50></td>
		<td><input type="text" name="mcashamt" value="<%=Mcashamt%>" size="10" ></td>
               
	</tr>
       	<tr>
               <td width=30></td>
		<td class="b12" align="left">庫房還款股金</td>
		<td width=50></td>
		<td><input type="text" name="msaveamt" value="<%=msaveamt%>" size="10"  ></td>
	</tr>

<% end if %>

 
<% if xtarget <> "T"  then %>
    
    
	<tr>
               <td width=30></td>
		<td class="b12" align="left">銀行還款利息</td>
		<td width=50></td>
		<td><input type="text" name="intamt" value="<%=intamt%>" size="10"  ></td>
	</tr>
      
	<tr>
               <td width=50></td>
		<td class="b12" align="left">銀行還款本金</td>
		<td width=50></td>
		<td><input type="text" name="cashamt" value="<%=cashamt%>" size="10" ></td>
       </tr>
    
       	<tr>
               <td width=30></td>
		<td class="b12" align="left">銀行還款股金</td>
		<td width=50></td>
		<td><input type="text" name="saveamt" value="<%=saveamt%>" size="10"  ></td>
	</tr>
<% end if%>
       


<tr>

                <td width=12"></td>
     		<td><font size="2" >社員狀況</formt></td>
                <td width=21></td>
		<td>
                     
			<select name="mstatus">
                        <option></option>
			<option<%if mstatus="L" then response.write " selected" end if%>>呆帳</option>
			<option<%if mstatus="D" then response.write " selected" end if%>>冷戶</option>
                        <option<%if mstatus="V" then response.write " selected" end if%>> IVA </option>
			<option<%if mstatus="C" then response.write " selected" end if%>>退社</option>
                        <option<%if mstatus="x" then response.write " selected" end if%>>凍結</option>
			<option<%if mstatus="P" then response.write " selected" end if%>>去世</option>
			<option<%if mstatus="B" then response.write " selected" end if%>>破產</option>
			<option<%if mstatus="N" then response.write " selected" end if%>>正常</option>
                        <option<%if mstatus="J" then response.write " selected" end if%>>新戶</option>
                        <option<%if mstatus="T" then response.write " selected" end if%>>庫房</option>
			<option<%if mstatus="H" then response.write " selected" end if%>>暫停銀行</option>
			<option<%if mstatus="A" then response.write " selected" end if%>>自動轉帳(ALL)</option>
			<option<%if mstatus="0" then response.write " selected" end if%>>自動轉帳(股金)</option>
			<option<%if mstatus="1" then response.write " selected" end if%>>自動轉帳(股金,利息)</option>
			<option<%if mstatus="2" then response.write " selected" end if%>>自動轉帳(股金,本金)</option>
			<option<%if mstatus="3" then response.write " selected" end if%>>自動轉帳(利息,本金)</option>
			<option<%if mstatus="M" then response.write " selected" end if%>>庫房,銀行</option>
			<option<%if mstatus="F" then response.write " selected" end if%>>特別個案</option>
                        <option<%if mstatus="8" then response.write " selected" end if%>>終止社籍轉帳</option>
                        <option<%if mstatus="9" then response.write " selected" end if%>>終止社籍正常</option>
			</select>
                 
		</td>
       </tr>
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if id <> "" then %>
               
		<%if session("userLevel")<>1 and session("userLevel")<>2 and session("userLevel")<>6 then%>
		<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
		<input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
                <input type="submit" value="刪除" onclick="return validating()&&confirm('確定刪除?')" name="delrec" class="sbttn">
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
				<td><input type="text" name="lnnum" value="<%=lnnum%>" size="12"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
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



        </table>
        </td>  
   </tr>
</table>       
</center>
</form>
</body>
</html>
