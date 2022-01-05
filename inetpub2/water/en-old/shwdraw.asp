<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "nloan.asp"'
end if
if request.form("calc") <> "" then
    
   monthrepay = request.form("appamt")/request.form("install")
end if
uid = request("uid")

if request.form("action") <> "" then
        uid = session("uid") 
        addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next
    set rs1 = server.createobject("ADODB.Recordset")
	msg = ""

	if Lnnum<>"" then
		sql1 = "select * from loanrec where lnnum=" & lnnum
		rs1.open sql1, conn,1
		if not rs1.eof then
			msg = "貸款號碼已經存在 "
		end if
		rs1.close
        else
  	  msg = "貸款號碼未填入 "
		
	end if
	if msg="" then
		conn.begintrans

conn.execute("update loanapp set deleted = 1 where uid='"&uid&"'   " )

	        sql1 = "select * from loanrec where 0=1"
		rs1.open sql1, conn, 3, 3
		rs1.addnew
                rs1("lnnum")=lnnum
                rs1("memno")=memno
             	if lndate<>"" then rs1("lndate") = right(lndate,4)&"/"&mid(lndate,4,2)&"/"&left(lndate,2) else rs1("lndate")="" end if
                rs1("appamt")=appamt
                rs1("monthrepay")=monthrepay
		rs1("install")=install
		IF autopamt<>"" then rs1("autopamt")=autopamt else rs1("autopamt")= 0 end if
		if rs1("salarydeduct")<>"" then rs1("salarydeduct")=salarydeduct else rs1("salarydeduct")=salarydeduct end if
		if intchoice="銀行" then
                   rs1("intchoice") ="A"
                else
                  rs1("intchoice") ="S"
                end if 
		rs("repaystat") = " "
		rs1.update
		
               
conn.execute("insert into memTx (memTxNo,memNo,txDate,treNo,amtloan,txAmt,lnnum,deleted) select max(memTxNo)+1,'"&rs1(0)&"','"&rs1(2)&"','0D',"&rs1(4)&","&rs1(4)&",'"&rs1(1)&"',0 from memTx")           


		rs1.close 	
         
        
		conn.committrans
		msg = "紀錄已更新"
		response.redirect( "nloan.asp")
	end if
else
	if uid <> "" then
		sql = "select * from loanapp where uid= " & uid
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "nloan.asp"
		else
			For Each Field in rs.fields		
			TheString = Field.name & "= rs(""" & Field.name & """)"
			Execute(TheString)
			Next
		end if
                install = rs("installment")
                monthrepay = rs("monthrepay")
                appamt = rs("loanamt")
                rs.close        
         end if  
               
		lndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
		session("uid") = uid		
end if
%>
<html>
<head>
<title>貸款資料修正</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">

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

	if (formObj.lnnum.value==""){
		reqField=reqField+", 貸款編號";
		if (!placeFocus)
			placeFocus=formObj.lnnum;
	}

	if (formObj.memNo.value==""){
		reqField=reqField+", 社員編號";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}

	if (formObj.reqamt.value==""){
		reqField=reqField+", 貸款金額";
		if (!placeFocus)
			placeFocus=formObj.reqamt;
	}

	if (formObj.appamt.value==""){
		reqField=reqField+", 批核金額";
		if (!placeFocus)
			placeFocus=formObj.appamt;
	}

	if (formObj.monthrepay.value==""){
		reqField=reqField+", ’每月還款";
		if (!placeFocus)
			placeFocus=formObj.monthrepay;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.lnnum.focus()">
<DIV>


<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>

<br>
<form name="form1" method="post" action="nloanDetail.asp">

<input type="hidden" name="id" value="<%=id%>">
<div><center><font size="3">新貸款建立</font></center></div>


<div style="z-index: 97; left: 350px; width: 500px; position: absolute; top: 160px;  height: 68px ">
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款號碼</td>
		<td width=50></td>
		<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"></td>
		<td><input type="submit" value="搜尋" name="Search" class="sbttn"></td>
    	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 183px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員號碼</td>
		<td width=50></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"></td> 
		<td width=50></td>
     
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 206px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">設定日期</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"</td>
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 229px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"</td>
	</tr>
</TABLE>
</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 252px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤分期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"</td>
	</tr>
</TABLE>
</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 275px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"</td>
	</tr>
</TABLE>
</div>

</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 298px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">自動轉賬</td>
		<td width=50></td>
		<td><input type="text" name="autopamt" value="<%=autopamt%>" size="10"></td>
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 321px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">庫房扣薪</td>
		<td width=50></td>
		<td><input type="text" name="salaryDedut" value="<%=salaryDedut%>" size="10"></td>
	</tr>
</TABLE>
</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 344px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">扣除利息</td>
		<td width=50></td>
		<td>
			<select name="intchoice">
			<option<%if intchoice="A" then response.write " selected" end if%>>銀行</option>
			<option<%if IntChoice="S" then response.write " selected" end if%>>庫房</option>
			</select>
		</td>
	</tr>
</TABLE>
</div>

<div style="z-index: 118; left: 350px; width: 1000px; position: absolute; top: 370px;  height: 68px">
<table id-"table20" border="0" cellspacing="0" cellpadding="0">       
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</DIV>
</div>

</form>
</body>
</html>
