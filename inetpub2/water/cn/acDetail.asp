<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "ac.asp"
end if

id = request("id")
if id="" then
	response.redirect "ac.asp"
end if
set rs = server.createobject("ADODB.Recordset")

sql = "select * from glControl"
rs.open sql, conn
acPeriod=rs("acPeriod")
rs.close

sql = "select * from memMaster where memNo=" & id
rs.open sql, conn, 2, 2
if rs.eof then
	response.redirect "ac.asp"
end if

if request.form("action") = "Save" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
	rs("salaryDedut") = salaryDedut
	rs("autopayAmt") = autopayAmt
	rs("autopayPerm") = autopayPerm
	rs("bankAccNo") = bankAccNo
	rs("leaguedue") = leaguedue
	rs.update
	addUserLog "Modify Account Detail"
	msg = "紀錄已更新"
end if

if id <> "" then
	For Each Field in rs.fields
		TheString = Field.name & "= rs(""" & Field.name & """)"
		Execute(TheString)
	Next
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>個人賬修正</title>
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.id.select();form1.id.focus();">
<!-- #include file="menu.asp" -->
<div align="right"><a href="memberDetail.asp?id=<%=request("id")%>">社員資料修正</a>&nbsp;&nbsp;</div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="acDetail.asp">
<table border="0" cellspacing="0" cellpadding="0">
	<tr height="40" valign="top">
		<td colspan="3" class="b8" align="center">
			社員號碼 <input type="text" name="id" value="<%=id%>" size="4">
			<input type="submit" value="選擇" name="Search" class="sbttn" onclick="return checkId()"> <%=memName%>
		</td>
	</tr>
	<tr>
		<td valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="b8" align="right">自動轉賬</td>
					<td width="10"></td>
					<td><input type="text" name="autopayAmt" value="<%=autopayAmt%>" size="25"></td>
				</tr>
				<tr>
					<td class="b8" align="right">預設轉賬額</td>
					<td></td>
					<td><input type="text" name="autopayPerm" value="<%=autopayPerm%>" size="25"></td>
				</tr>
				<tr>
					<td class="b8" align="right">庫房扣薪</td>
					<td></td>
					<td><input type="text" name="salaryDedut" value="<%=salaryDedut%>" size="25"></td>
				</tr>
				<tr>
					<td class="b8" align="right">每月還款額</td>
					<td></td>
					<td align="right" height="23"><%=formatnumber(loanRepaid,2)%></td>
				</tr>
				<tr>
					<td class="b8" align="right">平均利息</td>
					<td></td>
					<td align="right" height="23"><%=formatnumber(OSInterest,2)%></td>
				</tr>
			</table>
		</td>
		<td width=10></td>
		<td valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="b8" align="right">本月利息</td>
					<td width="10"></td>
					<td align="right" height="23"><%=formatnumber(thisInterest,2)%></td>
				</tr>
				<tr>
					<td class="b8" align="right">下月利息</td>
					<td></td>
					<td align="right" height="23"><%=formatnumber(eval("thisLoanBal"&acPeriod)*0.01-OSinterest+thisInterest,2) %></td>
				</tr>
				<tr>
					<td class="b8" align="right">銀行戶口</td>
					<td></td>
					<td><input type="text" name="bankAccNo" value="<%=bankAccNo%>" size="25" maxlength="50"></td>
				</tr>
				<tr>
					<td class="b8" align="right">扣除會費</td>
					<td></td>
					<td>
						<select name="leaguedue">
						<option value="-1"<%if leaguedue then response.write " selected" end if%>>Yes</option>
						<option value="0"<%if not leaguedue then response.write " selected" end if%>>No</option>
						</select>
					</td>
				</tr>
				<tr>
					<td class="b8" align="right">股息</td>
					<td></td>
					<td align="right" height="23"><%=formatnumber(dividend,2)%></td>
				</tr>
				<tr>
					<td colspan="3" align="right">
						<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
						<input type="submit" value="確定" onclick="return validating()&&confirm('Are you going to save the Record?')" name="action" class="sbttn">
						<%end if%>
						<input type="submit" value="返回" name="back" class="sbttn">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td rowspan="2"><font size="2" color="#FFFFFF">會計月</font></td>
	<td colspan="2"><font size="2" color="#FFFFFF">去年</font></td>
	<td colspan="2"><font size="2" color="#FFFFFF">今年</font></td>
  </tr>
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">存款金額</font></td>
	<td><font size="2" color="#FFFFFF">貸款金額</font></td>
	<td><font size="2" color="#FFFFFF">存款金額</font></td>
	<td><font size="2" color="#FFFFFF">貸款金額</font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">1 (Sep)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal1,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal1,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal1,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal1,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">2 (Oct)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal2,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal2,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal2,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal2,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">3 (Nov)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal3,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal3,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal3,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal3,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">4 (Dec)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal4,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal4,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal4,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal4,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">5 (Jan)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal5,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal5,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal5,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal5,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">6 (Feb)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal6,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal6,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal6,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal6,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">7 (Mar)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal7,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal7,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal7,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal7,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">8 (Apr)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal8,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal8,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal8,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal8,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">9 (May)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal9,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal9,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal9,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal9,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">10 (Jun)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal10,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal10,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal10,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal10,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">11 (Jul)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal11,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal11,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal11,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal11,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">12 (Aug)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastShrBal12,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(lastLoanBal12,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisShrBal12,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisLoanBal12,2)%></font></td>
  </tr>
</table>
</center>
</form>
</body>
</html>
