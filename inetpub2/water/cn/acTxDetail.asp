<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next

bdNum=cdbl(bdNum)
sql = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close
if acPeriod<=4 then
	m=acPeriod+8
	y=acYear
else
	m=acPeriod-4
	y=acYear+1
end if

if request("from") = Request.ServerVariables("script_name") then
	if request.form("Back")<>"" then
		response.redirect "acTx.asp"
	end if

	if msg = "" then
		if request.form("Add")<>"" then
			mDate = y&"/"&m&"/"&mDay

			conn.begintrans
			sql = "select top 1 * from memTx order by memTxNo desc"
			rs.open sql, conn, 2, 2
			if rs.eof then
				id = 1
			else
				id = rs("memTxNo")+1
			end if
			rs.addnew
			For Each Field in rs.fields
				if Field.name="memTxNo" then
					rs("memTxNo")=id
				elseif Field.name="txDate" then
					rs("txDate") = mDate
				elseif Field.name="treNo" then
					rs("treNo") = treNo
				else
					if request(Field.name)<>"" then
						TheString = "rs(""" & Field.name & """) = cdbl(" & Field.name & ")"
						Execute(TheString)
					end if
				end if
			Next
			rs.update
			rs.close

			if loanPaid<>"" then
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&loanPaid&" where glId='0205'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Load Paid','"&mDate&"','D',"&loanPaid&",0 from glTx")
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&loanPaid&" where glId='0201'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Load Paid','"&mDate&"','C',"&loanPaid&",0 from glTx")
				conn.execute("update memMaster set thisLoanBal"&acPeriod&"=thisLoanBal"&acPeriod&" - "&loanPaid&",CalcInterest="&CalcInterest&" where memNo="&memNo)
				conn.execute("update memMaster set overdue=0 where memNo="&memNo)
				Set rs1 = conn.execute("Select thisLoanBal"&acPeriod&" from memMaster where memNo="&memNo)
				if rs1(0) = cdbl(loanPaid) then
					conn.execute("update memMaster set thisInterest=0,OSinterest=0 where memNo="&memNo)
				end if
			end if
			if sharePaid<>"" then
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&sharePaid&" where glId='0205'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Share Paid','"&mDate&"','D',"&sharePaid&",0 from glTx")
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&sharePaid&" where glId='0401'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0401','Share Paid','"&mDate&"','C',"&sharePaid&",0 from glTx")
				conn.execute("update memMaster set thisShrBal"&acPeriod&"=thisShrBal"&acPeriod&" + "&sharePaid&" where memNo="&memNo)
			end if
			if shareWithdrawn<>"" then
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&shareWithdrawn&" where glId='0205'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Share Withdrawn','"&mDate&"','C',"&shareWithdrawn&",0 from glTx")
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&shareWithdrawn&" where glId='0401'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0401','Share Withdrawn','"&mDate&"','D',"&shareWithdrawn&",0 from glTx")
				conn.execute("update memMaster set thisShrBal"&acPeriod&"=thisShrBal"&acPeriod&" - "&shareWithdrawn&" where memNo="&memNo)
			end if
			if amtLoan<>"" then
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&amtLoan&" where glId='0205'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Amount Loaned','"&mDate&"','C',"&amtLoan&",0 from glTx")
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&amtLoan&" where glId='0201'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0201','Amount Loaned','"&mDate&"','D',"&amtLoan&",0 from glTx")
				conn.execute("update memMaster set thisLoanBal"&acPeriod&"=thisLoanBal"&acPeriod&" + "&amtLoan&" where memNo="&memNo)
			end if
			if interestPaid<>"" then
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&interestPaid&" where glId='0205'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0205','Interest Paid','"&mDate&"','D',"&interestPaid&",0 from glTx")
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&interestPaid&" where glId='0501'")
				conn.execute("insert into  glTx (glTxNo,glId,txItem,txDate,txType,txAmt,deleted) select max(glTxNo)+1,'0501','Interest Paid','"&mDate&"','C',"&interestPaid&",0 from glTx")
			end if
			addUserLog "Add Account transaction"
			conn.committrans

			bdNum=bdNum+1
			thisString="bd1c"&bdNum&"="""&right("0"&mDay,2)&"/"&right("0"&m,2)&"/"&y&""""
			execute(thisString)
			thisString="bd2c"&bdNum&"="""&treNo&""""
			execute(thisString)
			thisString="bd3c"&bdNum&"="""&sharePaid&""""
			execute(thisString)
			thisString="bd4c"&bdNum&"="""&shareWithdrawn&""""
			execute(thisString)
			thisString="bd5c"&bdNum&"="""&amtLoan&""""
			execute(thisString)
			thisString="bd6c"&bdNum&"="""&CalcInterest&""""
			execute(thisString)
			thisString="bd7c"&bdNum&"="""&monthlyRepaid&""""
			execute(thisString)
			thisString="bd8c"&bdNum&"="""&interestPaid&""""
			execute(thisString)
			thisString="bd9c"&bdNum&"="""&loanPaid&""""
			execute(thisString)
			treNo=""
			sharePaid=""
			shareWithdrawn=""
			amtLoan=""
			CalcInterest=-1
			monthlyRepaid=""
			interestPaid=""
			loanPaid=""
		end if
	end if

	SQL = "select memNo,memName,thisShrBal"&acPeriod&" as shareBal,thisLoanBal"&acPeriod&" as loanBal from memMaster where deleted=0 and memNo="&memNo
	rs.open sql, conn
	if not rs.eof then
		memName=rs("memName")
		shareBal=formatNumber(rs("shareBal"),2)
		loanBal=formatNumber(rs("loanBal"),2)
		msg = ""
	else
		msg="找不到社員號碼"
	end if
	rs.close

else
    CalcInterest=-1
    bdNum = 0
    shareBal="0.00"
    loanBal="0.00"
end if
%>
<html>
<head>
<title>個人賬入數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

function checkDay(mDay){
  D=mDay.value;
  M=<%=m%>;
  Y=<%=y%>;
  if(isNaN(D) || D=="")
    return false;
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
  return true;
}

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.memNo.value==""){
		reqField=reqField+", 社員號碼";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}

	if (formObj.mDay.value==""){
		reqField=reqField+", 日期";
		if (!placeFocus)
			placeFocus=formObj.mDay;
	}

	if (formObj.monthlyRepaid.value!=""&&formObj.amtLoan.value==""){
		reqField=reqField+", 貸款";
		if (!placeFocus)
			placeFocus=formObj.amtLoan;
	}

	if (formObj.amtLoan.value!=""&&formObj.monthlyRepaid.value==""){
		reqField=reqField+", 每月還款";
		if (!placeFocus)
			placeFocus=formObj.monthlyRepaid;
	}

	if (parseFloat(formObj.shareWithdrawn.value)>parseFloat(formObj.shareBal.value)){
		reqField=reqField+", 少於存款金額的退股金額";
		if (!placeFocus)
			placeFocus=formObj.shareWithdrawn;
	}

	if (formObj.sharePaid.value==""&&formObj.shareWithdrawn.value==""&&formObj.amtLoan.value==""&&formObj.monthlyRepaid.value==""&&formObj.interestPaid.value==""&&formObj.loanPaid.value==""){
		reqField=reqField+", 任何紀錄";
		if (!placeFocus)
			placeFocus=formObj.sharePaid;
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
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" bgcolor="#eeeef0" onload="form1.memNo.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form method="post" action="<%=request.servervariables("script_name")%>" name="form1">
<input type="hidden" name="From" value="<%=request.servervariables("script_name")%>">
<input type="hidden" name="shareBal" value="<%=shareBal%>">
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td class="b8" width="100">社員號碼 </td>
        <td width="100"><input type=text name=memNo value="<% =memNo %>" maxlength=4 size=4><input type="button" value="選擇" onclick="popup('pop_searchMem1.asp?key='+form1.memNo.value)" class="sbttn"></td>
        <td width="100"></td>
    </tr>
    <tr>
        <td class="b8" height="25">社員名稱 </td>
        <td colspan="2" id="memName"><%=memName%></td>
    </tr>
    <tr>
        <td class="b8" height="25">存款金額 </td>
        <td align="right" id="shareBal"><%=shareBal%></td>
        <td></td>
    </tr>
    <tr>
        <td class="b8" height="25">貸款金額 </td>
        <td align="right" id="loanBal"><%=loanBal%></td>
        <td></td>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="bottom" bgcolor="#87CEEB" height="17" align="center">
		<td class="n8" width=30>#</td>
		<td class="n8">日期</td>
		<td class="n8">類別</td>
		<td class="n8">存款</td>
		<td class="n8">退股</td>
		<td class="n8">貸款</td>
		<td class="n8">計息</td>
		<td class="n8">每月還款</td>
		<td class="n8">貸款利息</td>
		<td class="n8">還款</td>
		<td class="n8"></td>
	</tr>
<%
ttl3=0 : ttl4=0 : ttl5=0 : ttl8=0 : ttl9=0
for idx = 1 to bdNum%>
	<input type="hidden" name="bd1c<%=idx%>" value="<%=eval("bd1c"&idx)%>">
	<input type="hidden" name="bd2c<%=idx%>" value="<%=eval("bd2c"&idx)%>">
	<input type="hidden" name="bd3c<%=idx%>" value="<%=eval("bd3c"&idx)%>">
	<input type="hidden" name="bd4c<%=idx%>" value="<%=eval("bd4c"&idx)%>">
	<input type="hidden" name="bd5c<%=idx%>" value="<%=eval("bd5c"&idx)%>">
	<input type="hidden" name="bd6c<%=idx%>" value="<%=eval("bd6c"&idx)%>">
	<input type="hidden" name="bd7c<%=idx%>" value="<%=eval("bd7c"&idx)%>">
	<input type="hidden" name="bd8c<%=idx%>" value="<%=eval("bd8c"&idx)%>">
	<input type="hidden" name="bd9c<%=idx%>" value="<%=eval("bd9c"&idx)%>">
	<tr>
		<td class="n10" align="center"><%=idx%></td>
		<td class="show" align="center"><%=eval("bd1c"&idx)%></td>
		<td class="show" align="center"><%=eval("bd2c"&idx)%></td>
		<td class="show" align="right"><%if eval("bd3c"&idx)="" then response.write "&nbsp;" else response.write formatNumber(eval("bd3c"&idx),2) end if%></td>
		<td class="show" align="right"><%if eval("bd4c"&idx)="" then response.write "&nbsp;" else response.write formatNumber(eval("bd4c"&idx),2) end if%></td>
		<td class="show" align="right"><%if eval("bd5c"&idx)="" then response.write "&nbsp;" else response.write formatNumber(eval("bd5c"&idx),2) end if%></td>
		<td class="show" align="right"><%if eval("bd6c"&idx)<>0 then response.write "Yes" else response.write "No" end if%></td>
		<td class="show" align="right"><%if eval("bd7c"&idx)="" then response.write "&nbsp;" else response.write formatNumber(eval("bd7c"&idx),2) end if%></td>
		<td class="show" align="right"><%if eval("bd8c"&idx)="" then response.write "&nbsp;" else response.write formatNumber(eval("bd8c"&idx),2) end if%></td>
		<td class="show" align="right"><%if eval("bd9c"&idx)="" then response.write "&nbsp;" else response.write formatNumber(eval("bd9c"&idx),2) end if%></td>
		<td></td>
    </tr>
<%
	if eval("bd3c"&idx)<>"" then ttl3 = ttl3 + cdbl(eval("bd3c"&idx)) end if
	if eval("bd4c"&idx)<>"" then ttl4 = ttl4 + cdbl(eval("bd4c"&idx)) end if
	if eval("bd5c"&idx)<>"" then ttl5 = ttl5 + cdbl(eval("bd5c"&idx)) end if
	if eval("bd8c"&idx)<>"" then ttl8 = ttl8 + cdbl(eval("bd8c"&idx)) end if
	if eval("bd9c"&idx)<>"" then ttl9 = ttl9 + cdbl(eval("bd9c"&idx)) end if
next%>
	<input type="hidden" name="bdNum" value="<%=bdNum%>">
	<tr>
		<td align="center" class="n10"><%=idx%></td>
		<td><input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=m%>/<%=y%>&nbsp;</td>
		<td>
			<select name="treNo">
			  <option value="99"<%if treNo ="99" then%> selected<%end if%>>99</option>
			  <option value="AT"<%if treNo ="AT" then%> selected<%end if%>>AT</option>
			  <option value="SD"<%if treNo ="SD" then%> selected<%end if%>>SD</option>
			  <option value="AD"<%if treNo ="AD" then%> selected<%end if%>>AD</option>
			</select>
		</td>
		<td><input type=text name="sharePaid" value="<%=sharePaid%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type=text name="shareWithdrawn" value="<%=shareWithdrawn%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type=text name="amtLoan" value="<%=amtLoan%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td>
			<select name="calcInterest">
			<option value="-1"<% if calcInterest<>0 then %> selected<% end if%>>Yes</option>
			<option value="0"<% if calcInterest=0 then %> selected<% end if%>>No</option>
			</select>
		</td>
		<td><input type=text name="monthlyRepaid" value="<%=monthlyRepaid%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type=text name="interestPaid" value="<%=interestPaid%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type=text name="loanPaid" value="<%=loanPaid%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
        <td><input type="submit" value="新增" name="add" class="xbttn" onclick="return validating()"></td>
    </tr>
	<tr>
		<td align="right" colspan="3" class="b10">Total :</td>
		<td align="right"><%=formatNumber(ttl3,2)%></td>
		<td align="right"><%=formatNumber(ttl4,2)%></td>
		<td align="right"><%=formatNumber(ttl5,2)%></td>
		<td></td>
		<td></td>
		<td align="right"><%=formatNumber(ttl8,2)%></td>
		<td align="right"><%=formatNumber(ttl9,2)%></td>
		<td></td>
    </tr>
	<tr>
        <td colspan="10" align="right"><input type="submit" value="返回" name="back" class="sbttn"></td>
    </tr>
</table>
</form>
</center></div>
</body>
</html>
<%
conn.close
set conn=nothing
%>