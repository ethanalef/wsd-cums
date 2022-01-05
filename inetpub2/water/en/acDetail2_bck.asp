<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->

<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if

if request.form("Search")<>"" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        id = memno
        set rs = conn.execute("select memno,memname,memcname,memGrade,employCond,membday,monthsave,monthssave,tpayamt,mstatus,remark from memmaster where memno='"&memno&"' ")
        if not rs.eof then
	   		For Each Field in rs.fields
			if Field.name="memBday" or Field.name="firstAppointDate" or Field.name="memDate" or Field.name="Wdate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
                select case mstatus
                       case "L"
                           xstatus= "呆帳"
                           repayamt = 0
                       case  "D"
                           xstatus="冷戶"
                       	repayamt = 0
                       case  "V"
                           xstatus= " IVA "
                          repayamt = 0
                       case  "C"
                             xstatus= "退社"
             		   repayamt = 0	
                       case  "B"
                             xstatus= "去世"
                           repayamt = 0
                       case  "P"
                            xstatus="破產"
                          repayamt = 0
                       case  "N"
                            xstatus= "正常"
                            repayamt = monthsave    
                      case  "J"
                            xstatus= "新戶"
                            repayamt = monthsave
                      case "H"
                          xstatus= "暫停銀行"
                          repayamt = monthsave
                       case  "A"
                            xstatus="自動轉帳"
			    repayamt = monthsave
                       case  "0"
                            xstatus="自動轉帳(股金)"                       
			    repayamt = monthsave	

                       case  "1"
                            xstatus="自動轉帳(股金,利息)"
                            repayamt = monthsave
                       case  "Z"
                            xstatus="自動轉帳(股金,本金)"
                            repayamt = monthsave
                       case  "3"
                            xstatus="自動轉帳(利息,本金)"
                            repayamt = monthsave

                       case  "M"
                           xstatus = "庫房,銀行"
                           repayamt = monthssave
                      case  "T"
                            xstatus= "庫房"
                            repayamt = monthssave   
                end select
           chkdate =right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)
           age = year(date()) -year(memBday)
           set rs1 = conn.execute("select * from loanrec where memno='"&memno&"' and repaystat='N' ")
           if not rs1.eof then   
 	   		For Each Field in rs1.fields
			if Field.name="lndate" then
					TheString = "if rs1(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs1(""" & Field.name & """)),2)&""/""&right(""0""&month(rs1(""" & Field.name & """)),2)&""/""&year(rs1(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs1(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next 
                        if monthrepay="" then
                           monthrepay =iif( int(appamt/install)=appamt/install,int(appamt/install),int(appamt/install)+1)
                        end if
                       
           else
                lnnum=""          
           end if    
           rs1.close
           ttlbal = 0
           set rs1 = conn.execute("select * from share where memno='"&memno&"'  ")
           do while not rs1.eof
              select case rs1("code")
                     case "0A","A1","A2","A3","C0","C1","C3" ,"B6" 
                          ttlbal = ttlbal + rs1("amount")
                      case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3" 
                           ttlbal = ttlbal - rs1("amount")
               end select
                      
           rs1.movenext
           loop
           rs1.close
           
           if lnnum <> "" then
                pamt = 0
                pint = 0       
		set rs1  = conn.execute("select *  from loan where memno='"& memno & "'  and code='EI' and pflag= 1 ")				
                do while  not rs1.eof 
                      pamt = pamt + rs1("bal")
                rs1.movenext
                loop             
                rs1.close
		set rs1 = conn.execute("select * from loan where memno ='"& memno & "'  and code= 'FI' and pflag = 1 ")
                do while not  rs1.eof 
                      pint = pint + rs1("bal")       
                rs1.movenext
                loop
                rs1.close
                select case mstatus
                       case "A"
                          if (bal-monthrepay )>0 then                    
                              repayamt = monthrepay+(bal*.01)+pint+pamt+monthsave   
                          else
	         	     repayamt = bal+(bal*.01)+pint+pamt+monthsave  	
                         end if
                      case "T"
                            repayamt = tpayamt
                      case "M"
                                  
                         if (bal-monthrepay )>0 then                    
                              repayamt = monthrepay+(bal*.01)+pint+pamt+monthsave   
                          else
	         	     repayamt = bal+(bal*.01)+pint+pamt+monthsave  	
                         end if
                      case else
                           repayamt = 0 
                end select      
              set rs1=conn.execute("select * from guarantor where lnnum='"&lnnum&"' ")
              xx = 1
              do while not rs1.eof
                 select case xx
                        case  1 
                             guid1 = rs1("guarantorID")
                             guname1 = rs1("guarantorname")
                             gucname1 = rs1("guarantorcname")
                             gattlbal = 0
                             set rs2 = conn.execute("select * from share where memno='"&guid1&"'  ")
			     do while not rs2.eof
		              select case left(rs2("code"),1)
                		     case "A","0","C"
                         		 gattlbal = gattlbal + rs2("amount")
		                     case "B","G","H"
                		           gattlbal = gattlbal - rs2("amount")
		               end select
                      
           			rs2.movenext
		            loop
		            rs2.close
                            gusave1 = gattlbal
                       case  2
                             guid2 = rs1("guarantorID")
                             guname2 = rs1("guarantorname")
                             gucname2 = rs1("guarantorcname")
                             gattlbal = 0
                             set rs2 = conn.execute("select * from share where memno='"&guid2&"'  ")
			     do while not rs2.eof
		              select case left(rs2("code"),1)
                		     case "A","0","C"
                         		 gattlbal = gattlbal + rs2("amount")
		                     case "B","G","H"
                		           gattlbal = gattlbal - rs2("amount")
		               end select
                      
           			rs2.movenext
		            loop
		            rs2.close
                            gusave2 = gattlbal
                       case 3             
                             guid3 = rs1("guarantorID")
                             guname3 = rs1("guarantorname")
                             gucname3 = rs1("guarantorcname")
                             gattlbal = 0
                             set rs2 = conn.execute("select * from share where memno='"&guid3&"'  ")
			     do while not rs2.eof
		              select case left(rs2("code"),1)
                		     case "A","0","C"
                         		 gattlbal = gattlbal + rs2("amount")
		                     case "B","G","H"
                		           gattlbal = gattlbal - rs2("amount")
		               end select
                      
           			rs2.movenext
		            loop
		            rs2.close
                            gusave3 = gattlbal
                   end select
                    xx = xx + 1
              rs1.movenext
              loop 
              rs1.close   
          else
             
             

          end if
          set rs1=conn.execute("select * from guarantor where guarantorid='"&memno&"' ")
          xx = 1
          do while not rs1.eof
             select case xx
                    case 1          
                        guoid1 = rs1("memno")
                        guoln1 = rs1("lnnum")
                        set rs2 = conn.execute("select memno,memname,memcname from memmaster where memno='"&guoid1&"' ")
                        if not rs2.eof then
                           guoname1 = rs2("memname")
                           guocname1 = rs2("memcname")
                           
                        end if
                        rs2.close
                    case 2
                        guoid2 = rs1("memno")
                        guoln2 = rs1("lnnum")
                        set rs2 = conn.execute("select memno,memname,memcname from memmaster where memno='"&guoid2&"' ")
                        if not rs2.eof then
                           guoname2 = rs2("memname")
                           guocname2 = rs2("memcname")
			   
                        end if
                        rs2.close
                    case 3    
			 guoid3 = rs1("memno")
		         guoln3 = rs1("lnnum")
                        set rs2 = conn.execute("select memno,memname,memcname from memmaster where memno='"&guoid3&"' ")
                        if not rs2.eof then
			  	
                           guoname3 = rs2("memname")
                           guocname3 = rs2("memcname")
			   
                        end if
                        rs2.close
             end select
                        guoid1 = rs1("memno")
                        guoln1 = rs1("lnnum")
             xx = xx + 1   
          rs1.movenext
          loop
          rs1.close      

                cleardate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())          
                todate    = cleardate
                cashint = 0
                cashamt = 0
                ttlpamt = 0
        else
             msg ="社員編號不存在!"  
             memno = ""  
             tpayamt = 0
             repayamt = 0               
        end if
        rs.close

else
   memno = ""
   memcName =""
   memName  = ""
   chkdate = ""
   saveamt  = 0.00
   id = ""   
end if
%>
<html>
<head>
<title>個人帳查詢</title>
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

  var strValidChars = "0123456789/";
  var strChar = "";

 
   for (i = 0; i < dSize ; i++)
      {  
      strChar = cDate.substr(i,1);
      if ( strValidChars.indexOf(strChar) == -1)
         { 
         return false ;   
          }
      }

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



function checkId(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (!formatNum(formObj.id)){
        alert("Please fill correct account No.");
		form1.id.select();form1.id.focus();
        return false;
    }else{
        return true;
    }
}

function calculation(){
	formObj=document.form1;

 
	if (formObj.cleardate.value!=""){
           
            if (formObj.pint.value!=""){
               ppint1 = parseInt(formObj.pint.value)
            }else{
               ppint1 = 0
            }      
	    Payamt = parseInt(formObj.bal.value)    

           ssdate = formObj.cleardate.value
           ttdate = formObj.todate.value 
           if (ssdate >= ttdate){
	   Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31)

	   Y=ssdate.substr(6,4)           
           M=ssdate.substr(3,2)
	   D=ssdate.substr(0,2)
	   YY=parseInt(ssdate.substr(6,4)) 
           MM=parseInt(ssdate.substr(3,2))
	   DD=parseInt(ssdate.substr(0,2))
           chkdate1 = Y+M

           TY=ttdate.substr(6,4) 
           TM=ttdate.substr(3,2)
           TD=ttdate.substr(0,2)
           YYY=parseInt(ttdate.substr(6,4)) 
           YMM=parseInt(ttdate.substr(3,2))
           YDD=parseInt(ttdate.substr(0,2))
           chkdate2 = TY+TM
	   DM = 0
     
           if (chkdate1 < chkdate2){
              DM = -1  
           }else{
             if ( YY > YYY){
                DM = (YY-YYY)*12+ MM-1
             }else{
                if ( MM > YMM)
                   DM = MM - YMM
             }                     
           } 
             
	      mD=Months[M-1]
              ppint3 =Math.round( Payamt*0.01*DD/mD )
              ttlint = ppint1+ppint3
              if (DM > 0 ){
                
                 ppint2 =Math.round( Payamt*0.01*DM )
                 ttlint = ppint1 + ppint2 + ppint3  
              }
              document.form1.cashamt.value =Payamt
              document.all.tags( "td" )['cashamt'].innerHTML  = Payamt
              document.all.tags( "td" )['cashint'].innerHTML = ttlint
              document.all.tags( "td" )['ttlpamt'].innerHTML = Payamt+ttlint
             }else{
              document.all.tags( "td" )['cashamt'].innerHTML = 0
              document.all.tags( "td" )['cashint'].innerHTML = 0
              document.all.tags( "td" )['ttlpamt'].innerHTML = 0
             }           
	}else{
         
              document.all.tags( "td" )['cashamt'].innerHTML = 0
              document.all.tags( "td" )['cashint'].innerHTML = 0
              document.all.tags( "td" )['ttlpamt'].innerHTML = 0
         }
}


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (!formatNum(formObj.salaryDedut)){
		reqField=reqField+", 庫房扣薪";
		if (!placeFocus)
			placeFocus=formObj.salaryDedut;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="calculation();form1.memNo.focus()">
<DIV>
<!-- #include file="menu.asp" -->
<div align="right"><a href="membermod2.asp?id=<%=request("id")%>">社員資料修正</a>&nbsp;&nbsp;</div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
</center>
<br>

<div><center><font size="3">個人帳查詢</font></center></div>
<form name="form1" method="post" action="acDetail2.asp">
<table border="0" cellspacing="0" cellpadding="0">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memcName" value="<%=memcName%>">
<input type="hidden" name="memGrade" value="<%=memGrade%>">
<input type="hidden" name="age" value="<%=age%>">
<input type="hidden" name="employCond" value="<%=employCond%>">
<input type="hidden" name="firstAppointDate" value="<%=firstAppointDate%>">
<input type="hidden" name="lnnum" value="<%=lnnum%>">
<input type="hidden" name="bal" value="<%=bal%>">
<input type="hidden" name="appamt" value="<%=appamt%>">
<input type="hidden" name="lndate" value="<%=lndate%>">
<input type="hidden" name="monthsave" value="<%=monthsave%>">
<input type="hidden" name="monthssave" value="<%=monthssave%>">
<input type="hidden" name="tpayamt" value="<%=tpayamt%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="pamt" value="<%=pamt%>">
<input type="hidden" name="pint" value="<%=pint%>">
<input type="hidden" name="cashamt" value="<%=cashamt%>">
<input type="hidden" name="cashint" value="<%=cashint%>">
<input type="hidden" name="ttlpamt" value="<%=ttlpamt%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="repayamt" value="<%=repayamt%>">
<input type="hidden" name="remark" value="<%=remark%>">
<input type="hidden" name="repaystatus" value="<%=repaystatus%>">
	<tr>
		<td width="300" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
                                 
				<tr>
                                        <td width=10></td>
					<td class="b8" align="right">社員編號</td>
					<td width=10></td>
					<td>
					<input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"  <%if id<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>>
					<%if id = "" then %>
					<input type="button" value="選擇"  onclick="popup('pop_srhMemnoM.asp')" class="sbttn"  >
					<input type="submit" value="搜尋" name="Search" class ="Sbttn">
					<%else%>
			                <input type="submit" value="取消" name="bye" class="sbttn">
					<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>                        
							<%end if%>
					<% end if %>
					</td>
					<td><input type="submit" value="返回" name="back" class="sbttn"></td>

	
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">姓名</td>
					<td width=10></td>
					<td id="memName"><%=memName%></td>
					<td id="memcName"><%=memcName%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">職位</td>
					<td width=10></td>
					<td id="memGrade"><%=memGrade%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">年齡</td>
					<td width=10></td>
					<td id="age"><%=age%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">招聘條款</td>
					<td width=10></td>
					<td id="employCond"><%=employCond%></td>
                                </TR>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">狀況</td>
					<td width=10></td>
					<td id="xstatus"><%=xstatus%></td>
                                </TR>
                               </tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">備註</td>
					<td width=10></td>
					<td id="remark"><font size="2" color="red"><%=remark%></font></td>
                                </TR>
			</table>	
		    </td>
  		<td width="350" valign="top">
		<table border="0" cellspacing="0" cellpadding="0">                  
 		<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">1.擔保人</td>
			<td width=10></td>			
			<td id="guid1"><%=guid1%></td>
			<td width=10></td>
                        <td id="guname1"><%=guname1%></td>
			<td id="gucname1"><%=gucname1%></td>
		</tr>
                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">儲蓄結餘</td>
			<td width=10></td>  
 			<% if gusave1 > 0 then  %>                     
                        <td id="gusave1"><%=formatNumber(gusave1,2)%></td>                        
                        <%end if %> 
                        </tr>
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">2.擔保人</td>
			<td width=10></td>
			<td id="guid2"><%=guid2%></td>
			<td width=10></td>
                        <td id="guname2"><%=guname2%></td>
			<td id="gucname2"><%=gucname2%></td>
			</tr>
                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">儲蓄結餘</td>
			<td width=10></td>   
                        <%if gusave2 > 0 then %>                     
                        <td id="gusave2"><%=formatNumber(gusave2,2)%></td>  
                        <% end if %> 
                        </tr>
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">3.擔保人</td>
			<td width=10></td>
			<td id="guid3"><%=guid3%></td>
                       <td width=10></td> 
                        <td id="guname3"><%=guname3%></td>
			<td id="gucname3"><%=gucname3%></td>
			</tr>
                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">儲蓄結餘</td>
			<td width=10></td>   
                        <% if gusave3 > 0 then %>                   
                        <td id="gusave3"><%=formatNumber(gusave3,2)%></td>   
                        <% end if %>
                        </tr> 

                </table>
                </td>
  		<td width="400" valign="top">
		<table border="0" cellspacing="0" cellpadding="0">    
<%if guoid1<>"" then %>              
 		<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">1.擔保其他人</td>
			<td width=5></td>			
			<td id="guoid1"><%=guoid1%></td>
                </tr>
                <tr >
			<td width=10></td>
			<td class="b8" align="right">姓名</td>
			<td width=5></td>
                        <td id="guoname1"><%=guoname1%></td>
                    
		</tr>

                 <tr height="22">
			
			<td width=10></td>	
			<td class="b8" align="right">貸款編號</td>
			<td width=10></td>  
 			<% if guold1 <> "" then  %>                     
                        <td id="guln1"><%=guln1%></td>                        
                        <%end if %> 
                        </tr>
<%end if%>
<%if guoid2<>"" then %>
			<tr height="22" >
			<td width=10></td>
			<td class="b8" align="right">2.擔保其他人</td>
			<td width=10></td>
			<td id="guoid2"><%=guoid2%></td>
                </tr>
                <tr >
			<td width=10></td>
			<td class="b8" align="right">姓名</td>
			<td width=5></td>
                        <td id="guoname2"><%=guoname2%></td>
                    
		</tr>

                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">貸款編號</td>
			<td width=10></td>   
                        <%if guoid2 <> "" then %>                     
                        <td id="guln2"><%=goln2%></td>  
                        <% end if %> 
                        </tr>
<%end if %>
<%if guoid3 <> "" then %>
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">3.擔保其他人</td>
			<td width=10></td>
			<td id="guoid3"><%=guoid3%></td>
			                </tr>
                <tr >
			<td width=10></td>
			<td class="b8" align="right">姓名</td>
			<td width=5></td>
                        <td id="guoname3"><%=guoname3%></td>
                    
		</tr>

                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">貸款編號</td>
			<td width=10></td>   
                        <% if guoid3 <> "" then %>                   
                        <td id="guln3"><%=guln3%></td>   
                        <% end if %>
                        </tr> 
<%end if %>
                </table>
                </td>
          </tr>

</table>	
</tr>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
		<td width="300" valign="top">
                 <table border="0" cellspacing="0" cellpadding="0">	
		<%if memno <> "" then %>
			<tr height="22">	
			<td class="b8" align="right">股金結餘</td>
			<td width=30></td>
			<td id="ttlbal"><b><%=formatNumber(ttlbal,2)%></b></td>
                        </tr>
                        <tr height="22">			
			<%if mstatus="A" or mstatus="M" then %>
			<td class="b8" align="right">自動轉帳</td>
			<td width=10></td>
			<td id="monthsave"><%=monthsave%></td>
			<td width=10></td>
                        </tr> 
		<%end if %>
         
		<%if mstatus="T" then %>
                        <tr>
         	        <td class="b8" align="right">庫房扣薪(股金)</td>
			<td width=30></td>
			<td  id="monthssave" ><%=formatnumber(monthssave,2)%></td>
			<%end if %>
		        </tr>
                <%end if %>
		<%if tpayamt <>0 then %>
                       <tr>
         	        <td class="b8" align="right">庫房扣薪</td>
			<td width=30></td>
                         
			<td id="tpayamt"><%=formatnumber(tpayamt,2)%></td>
			
		        </tr>
		<%end if %>
               </table>
               </td>

		<%if lnnum <> "" then %>	
		<td width="300" valign="top">
                        <table border="0" cellspacing="0" cellpadding="0"> 
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">貸款編號</td>
			<td width=10></td>
			<td id="lnnum"><%=lnnum%></td>
                        <td width=10></td>
                        </tr>

			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">期數</td>
			<td width=10></td>
			<td id="install"><%=install%></td>
                        </tr>
                        <tr height="22">
                        <td width=10></td> 
                        <td class="b8" align="right">每月還款</td>
			<td width=10></td> 
			<td id="monthrepay"><%=formatNumber(monthrepay,2)%></td>
                        </tr>

			
                        </table>
                        </td>
              
		<td width="300" valign="top">
                 <table border="0" cellspacing="0" cellpadding="0">                                 
			<tr height="22">
			<td width=10></td>                 
                        <td class="b8" align="right">取票日期</td>
			<td width=10></td> 
			<td id="lndate"><%=lndate%></td>
                        </tr>
		<tr height="22">
			<td width=10></td> 
                        <td class="b8" align="right">貸款金額</td>
			<td width=10></td> 
			<td id="appamt"><%=formatNumber(appamt,2)%></td>
			</tr>    

 			<tr height="22">	
			<td width=10></td> 
                        <td class="b8" align="right">貸款結餘</td>
			<td width=10></td> 
			<td id="bal"><font color=red><%=formatNumber(bal,2)%></font></td>
			</tr>       
                   </table>
                   </td>    
	
<% end if %>
</tr>
</table>
<table

<%if repaystat="N" then %>
				<tr>
					<td width=10></td>
					<td class="b8" align="right">清數日期</td>
					<td width=10></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10" onblur="if(!formatDate(this)){this.value=''};calculation();"></td>
				</tr>
	<tr>
               <td width=10></td>
		<td class="b8" align="right">清數本金</td>
		<td width=10></td>
		<td id="cashamt"><%=formatNumber(cashamt,2)%></td>	


               <td width=10></td>
		<td class="b8" align="right">清數利息</td>
		<td width=10></td>
		<td id="cashint"><%=formatNumber(cashint,2)%></td>	

               <td width=10></td>
		<td class="b8" align="right">清數金額合共</td>
		<td width=10></td>
		<td id="ttlpamt"><%=formatNumber(ttlpamt,2)%></td>
	
	</tr>  
<%end if%>
				<tr>
					<td width=10></td>
					<td class="b8" align="right">瀏覽日期</td>
					<td width=10></td>
					<td><input type="text" name="chkdate" value="<%=chkdate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
				</tr>

</table>
<td colspan="3" align="right">


<% if id <> "" then %>
		<input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
<% end if %>
		</td>
	</tr>
</table>




</form>
</body>
</html>

 