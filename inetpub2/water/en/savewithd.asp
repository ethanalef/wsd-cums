<!-- #include file="../conn.asp" -->
<!-- #include file="cutpro.asp" -->
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
   work = 0
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
   if id ="" then          
        set rs = conn.execute("select memno,memname,memcname,mstatus from memmaster where memno='"&memno&"'  ")
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
          case "B"
                id = "" 
		msg = "** 社員巳破產 **"
          case "P"
                id =""
		msg = "** 社員巳去世 **"
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
   else
      work = 1         
   end if
   rs.close
   ttlsave = 0
   bal = 0
   appamt = 0
 
   if work = 1 then
   
      set rs= conn.execute("select * from share where memno='"&memno&"' order by memno,ldate,code  ")
      do while not rs.eof 
   
         select case rs("code")
                 case "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" ,"C5","0A","A8"
                    ttlsave = ttlsave + rs("amount")
                case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3","MF","B8"
                     ttlsave = ttlsave - rs("amount")
         end select       
      rs.movenext
      loop
      rs.close  
      set rs = conn.execute("select lnnum,lndate,appamt,bal from loanrec where repaystat='N' and memno='"&memno&"' ")
      if not rs.eof then
			For Each Field in rs.fields
			if Field.name="lndate"   then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
      end if 
      rs.close
   end if

end if


if request.form("action") <>"" then

	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
  
 	conn.begintrans
        pamt = amount
        pdate = right(ldate,4)&"/" &mid(ldate,4,2)&"/"&left(ldate,2)
        if lnflag="Y"  then
            conn.execute("insert into share (memno,code,ldate,amount,lnflag,sdesc) values ('"&memno&"','B1','"&pdate&"',"&pamt&",'"&lnflag&"','"&sdesc&"' ) ")
        elseif chkNegAdj = "1"  then
            conn.execute("insert into share (memno,code,ldate,amount) values ('"&memno&"','B8','"&pdate&"',"&pamt&") ")
        else
              conn.execute("insert into share (memno,code,ldate,amount) values ('"&memno&"','B1','"&pdate&"',"&pamt&") ")
        end if  
        if ttlsave = 0 then
            conn.execute("update memmaster set wdate='"&pdate&"' where memno='"&memno&"' ")
         end if
	conn.committrans
        amount = 0
        memno=""
	addUserLog "add Withdraw Detail"
	msg = "紀錄已更新"
end if
       
       
        yy = year(date())-1
	ldate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
        todate= "01/"&right("0"&month(date()),2)&"/"&yy 
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>退股建立</title>
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


function clearln(){        
     formObj=document.form1;
     if (formObj.lnflag.checked ){       
         
     
          document.form1.sdesc.value =''
         document.form1.lnflag.value = 'Y'
} 
}

function onNegAdj() {
	var checked = document.form1.chkNegAdj.checked;
  if (checked) {
		document.form1.chkNegAdj.value = "1";
  }
  else {
		document.form1.chkNegAdj.value = "0";
  }
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if ((formObj.amount.value=='')){
		reqField=reqField+", 金額";
		if (!placeFocus)
			placeFocus=formObj.amount;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNo.select();form1.memNo.focus();">
<!-- #include file="menu.asp" -->
<div><center><font size="3">退股建立</font></center></div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="savewithd.asp">
<table border="0" cellspacing="0" cellpadding="0">
<input type="hidden" name="uid" value="<%=uid%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="ttlsave" value="<%=ttlsave%>">
<input type="hidden" name="lnnum" value="<%=lnnum%>">
<input type="hidden" name="appamt" value="<%=appamt%>">
<input type="hidden" name="bal" value="<%=bal%>">
<input type="hidden" name="lndate" value="<%=lndate%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<table border="0" cellspacing="0" cellpadding="0" align="center">
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
					<td id="memName"><%=memName%></td>
				</tr>

				<tr height="22">
					<td class="b8" align="right"></td>
					<td width=10></td>
					<td id="memcName"><%=memcName%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">日期</td>
					<td width=10></td>
					<td><input type="text" name="ldate" value="<%=ldate%>" size="10" maxlength="10"onblur="if(!formatDate(this)){this.value=''};form1.ldate.value=this.value">
				</tr>
				<tr height="22">
					<td class="b8" align="right">貸款編號</td>
					<td width=10></td>
					<td id="lnnum"><%=LNNUM%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">取票日期</td>
					<td width=10></td>
					<td id="lndate" ><%=lndate%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">貸款總額</td>
					<td width=10></td>
					<td width="40" align="right" id="appamt"><%=formatnumber(appamt,2)%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">尚欠貸款金額</td>
					<td width=10></td>
					<td width="40" align="right" id="bal"><%=formatnumber(bal,2)%></td>
				</tr>

				<tr height="22">
					<td class="b8" align="right">股金結餘</td>
					<td width=10></td>
					<td withd=="40" id="ttlsave"><%=formatnumber(ttlsave,2)%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">金額</td>
					<td width=10></td>
					<td><input type="text" name="amount" value="<%=amount%>" size="10" maxlength="10" ></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">調(-)整</td>
					<td width=10></td>
					<td><input type="checkbox" name="chkNegAdj" onClick="onNegAdj()">
				</tr>
	       <tr height="22">
               <td class="b8" align="right">退回回贈</td>
               <td width=10></td> 
               <td><input type="checkbox" name="lnflag"   value="<% =lnflag %>"<%if lnflag="Y" then response.write " checked" end if%> onclick="if (this.checked){this.value=''};form1.lnflag.value='Y';form1.sdesc.value=''  ">
	</tr>                                
		
	<tr height="22" >

               <td class="b12" align="right">備註</td>
               <td width=10></td>  
               <td><input type="text" name="sdesc" value="<%=sdesc%>" size="35" maxlength="35"></td>
        </tr>
		</td>
	</tr>
</table>
				<tr>
					<td colspan="3" align="right">
                                        <%if id <>"" then %>
						<%if session("userLevel")<>2 and session("userLevel")<>1 and session("userLevel")<>4 then%>
						<input type="submit" value="確定" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
						<%end if%>
                                         <%end if%> 
<%if id<>"" then %>     
						<input type="submit" value="查詢個人帳" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.todate.value)" class="sbttn">
<%end if %>
						<input type="submit" value="取消" name="clrSrc" class="sbttn">
						<input type="submit" value="返回" name="bye" class="sbttn">
				</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>

</center>
</form>
</body>
</html>
