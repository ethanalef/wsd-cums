<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "cheque.asp"
end if
if request.form("new") <> "" then
	response.redirect "chequeDetail.asp"
end if

id = request("id")

if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if id="" then
		sql = "select count(*) from cheque where chequeNum='" & chequeNum & "'"
		rs.open sql, conn
		if rs(0) > 0 then
			msg = "支票號碼已經存在 "
		end if
		rs.close
	end if

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select top 1 * from cheque order by uid desc"
		else
			sql = "select * from cheque where uid=" & id
		end if
		rs.open sql, conn, 2, 2
		if id = "" then
			if rs.eof then
				id = 1
			else
				id = rs("uid") + 1
			end if
			rs.addnew
			rs("uid") = id
			addUserLog "Add Cheque"
		else
			addUserLog "Modify Cheque Detail"
		end if
		if chequeDate<>"" then rs("chequeDate") = right(chequeDate,4)&"/"&mid(chequeDate,4,2)&"/"&left(chequeDate,2) else rs("chequeDate")="" end if
		rs("chequeNum") = chequeNum
		rs("amount") = amount
		rs("payee") = payee
		rs.update
		conn.committrans
		msg = "紀錄已更新"
	end if
else
	if id <> "" then
		sql = "select * from cheque where uid=" & id
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "cheque.asp"
		else
			For Each Field in rs.fields
				TheString = Field.name & "= rs(""" & Field.name & """)"
				Execute(TheString)
			Next
		end if
	else
		chequeDate=date()
	end if
	chequeDate = right("0"&day(chequeDate),2)&"/"&right("0"&month(chequeDate),2)&"/"&year(chequeDate)
end if
%>
<html>
<head>
<title>支票對數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
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

	if (formObj.chequeDate.value==""||!formatDate(formObj.chequeDate)){
		reqField=reqField+", 支票日期";
		if (!placeFocus)
			placeFocus=formObj.chequeDate;
	}

	if (formObj.chequeNum.value==""){
		reqField=reqField+", 支票號碼";
		if (!placeFocus)
			placeFocus=formObj.chequeNum;
	}

	if (formObj.amount.value==""||!formatNum(formObj.amount)){
		reqField=reqField+", 金額";
		if (!placeFocus)
			placeFocus=formObj.amount;
	}

	if (formObj.payee.value==""){
		reqField=reqField+", 收款人";
		if (!placeFocus)
			placeFocus=formObj.payee;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.chequeDate.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="chequeDetail.asp">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">支票日期 (dd/mm/yy)</td>
		<td width=10></td>
		<td><input type="text" name="chequeDate" value="<%=chequeDate%>" size="50" maxlength="10"></td>
	</tr>
	<tr>
		<td class="b8" align="right">支票號碼</td>
		<td width=10></td>
		<td><input type="text" name="chequeNum" value="<%=chequeNum%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">金額</td>
		<td width=10></td>
		<td><input type="text" name="amount" value="<%=amount%>" size="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">收款人</td>
		<td width=10></td>
		<td><input type="text" name="payee" value="<%=payee%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="確定" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<input type="submit" value="新增" name="new" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
