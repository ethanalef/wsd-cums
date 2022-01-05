<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "shwdloan.asp"'
end if
if request.form("calc") <> "" then
		For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
		Next
                set rs2 = server.createobject("ADODB.Recordset")
	        sql1 = "select * from memtx where treno like 'F%' and lnnum="&lnnum&" order by txdate desc"
		rs2.open sql1,conn,2,2
                xx = 0
                ointamt = 0 
                do while not rs2.eof and xx = 0
                   if rs2("treno") ="FI" then
                      ointamt = ointamt + rs2("interestpaid")
                      xx = 1
                   else
                      xx = 1
                   end if
                   rs2.movenext
                loop
                rs2.close
       y = year(cleardate)
       m = month(cleardate)
       d = day(cleardate)
       if y/100=int(y/100) and y/4=int(y/4) then
           daylist ="312931303130313130313031"
           nday = mid(daylist,(m-1)*2+1,2)
        else
           daylist ="312831303130313130313031"
           nday = mid(daylist,(m-1)*2+1,2)         
	end if         
        nintamt =round( bal * 0.01 * d / nday ,0)
       
         cashamt = ointamt+bal + nintamt    
end if
xlnnum = request("lnnum")

if request.form("action") <> "" then
        addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next
  	set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if cleardate="" then
           	msg = "清數日期未填入 "
        else
        if cashamt  = "" then
        
		msg = "清數金額未填入 "
        end if
        end if




	if msg="" then
		conn.begintrans                
               
                set rs = server.createobject("ADODB.Recordset")
	        sql = "select * from loanrec where LNNUM="&XLNNUM
		rs.open sql, conn, 2, 2
		
   
             	if cleardate<>"" then rs("cleardate") = right(cleardate,4)&"/"&mid(cleardate,4,2)&"/"&left(cleardate,2)  end if
                rs("repaystat")="C"

		rs.update
conn.execute("insert into memTx (memTxNo,memNo,txDate,treNo,shareWithdrawn,txAmt,lnnum,deleted) select max(memTxNo)+1,'"&rs(0)&"','"&rs(7)&"','B4',"&rs(9)&","&cashamt&",'"&cashamt&"',0 from memTx")           
		
              
		rs.close 	
         
        
		conn.committrans
		msg = "紀錄已更新"
             end if
	
else
	if xlnnum <> "" then
		sql = "select * from loanrec where lnnum= " & xlnnum
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "shwdloan.asp"
		else
			For Each Field in rs.fields
			if Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if
                rs.close 

	 end if
	cleardate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
end if
%>
<html>
<head>
<title>貸款清數建立</title>
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

	if (formObj.cleardate.value==""){
		reqField=reqField+", ’清數日期";
		if (!placeFocus)
			placeFocus=formObj.cleardate;
	}

	if (formObj.newlnnum.value==""){
		reqField=reqField+", ’新貸款號碼";
		if (!placeFocus)
			placeFocus=formObj.newlnnum;
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
<!-- #include file="menu.asp" -->

<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>

<br>
<form name="form1" method="post" action="shwdlndetail.asp">

<input type="hidden" name="id" value="<%=id%>">
<div><center><font size="3">股金清數建立</font></center></div>


<div style="z-index: 97; left: 350px; width: 236px; position: absolute; top: 160px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款號碼</td>
		<td width=50></td>
		<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 183px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員號碼</td>
		<td width=50></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
		
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 206px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">設定日期</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 229px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 252px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤分期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 275px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>

</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 298px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">本金結餘</td>
		<td width=50></td>
		<td><input type="text" name="bal" value="<%=bal%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 321px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數日期</td>
		<td width=50></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10"></td>
	</tr>
</TABLE>
</div>

<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 344px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數金額</td>
		<td width=50></td>
		<td><input type="text" name="cashamt" value="<%=cashamt%>" size="10"></td>
	</tr>
</TABLE>
</div>

<div style="z-index: 118; left: 350px; width: 1000px; position: absolute; top: 388px;  height: 68px">
<table id-"table20" border="0" cellspacing="0" cellpadding="0">       
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="計利息" name="calc" class="sbttn">
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</DIV>
</div>
</form>
</body>
</html>
