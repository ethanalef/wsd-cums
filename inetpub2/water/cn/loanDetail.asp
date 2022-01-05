<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "loan.asp"
end if
if request.form("new") <> "" then
	response.redirect "loanDetail.asp"
end if

uid = request("uid")

if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	conn.begintrans
	if uid = "" then
		sql = "select top 1 * from loanApp order by uid desc"
	else
		sql = "select * from loanApp where uid="&uid
	end if
	rs.open sql, conn, 2, 2
	if uid = "" then
		if rs.eof then
			uid = 1
		else
			uid = rs("uid")+1
		end if
		rs.addnew
		addUserLog "Add Loan Application"
	else
		addUserLog "Modify Loan Application"
	end if
	For Each Field in rs.fields
		if Field.name="deleted" then
		elseif Field.name="appDate" or Field.name="chequeDate" or Field.name="interviewDate" or Field.name="firstApprovalDate" or Field.name="secondApprovalDate" then
			TheString = "if " & Field.name & "<>"""" then rs(""" & Field.name & """) = right(" & Field.name & ",4)&""/""&mid(" & Field.name & ",4,2)&""/""&left(" & Field.name & ",2) else rs(""" & Field.name & """)=null end if"
			Execute(TheString)
		elseif Field.name="netSalary" or Field.name="loanAmt" or Field.name="installment" or Field.name="guarantorID" or Field.name="guarantorSalary" then
			if request(Field.name)<>"" then
				TheString = "rs(""" & Field.name & """) = cdbl(" & Field.name & ")"
				Execute(TheString)
			end if
		else
			TheString = "rs(""" & Field.name & """) = " & Field.name
			Execute(TheString)
		end if
	Next
	rs.update

	conn.execute("delete from loanReason where loanAppID="&uid)
	A = split(request("TS"),",",-1,1)
	if isarray(A) then
		if (ubound(A) >= 0) then
			for i = 0 to ubound(A)
				conn.execute("insert into loanReason (loanAppID,reasonID) values ("&uid&","&A(i)&")")
			next
		end if
	end if

	conn.committrans
'	msg = "紀錄已更新"
	rs.close
	set rs=nothing
	response.redirect "loanDetail.asp"
else
	if uid <> "" then
		sql = "select * from loanApp where uid="&uid
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "loan.asp"
		else
			For Each Field in rs.fields
				if Field.name="appDate" or Field.name="chequeDate" or Field.name="interviewDate" or Field.name="firstApprovalDate" or Field.name="secondApprovalDate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if
		rs.close
		set rs=nothing
	else
		appDate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
		chequeDate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
		interest=-1
		loanPlanID=0
		SpecialPlanID=0
	end if
end if
%>
<html>
<head>
<title>貸款申請</title>
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
		reqField=reqField+", 社員號碼";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}

	if (formObj.loanAmt.value==""||formObj.loanAmt.value==0){
		reqField=reqField+", 貸款額";
		if (!placeFocus)
			placeFocus=formObj.loanAmt;
	}

	if (formObj.installment.value==""||formObj.installment.value==0){
		reqField=reqField+", 攤分期數";
		if (!placeFocus)
			placeFocus=formObj.installment;
	}

	if (formObj.chequeDate.value==""){
		reqField=reqField+", 預定支票日期";
		if (!placeFocus)
			placeFocus=formObj.chequeDate;
	}

	var TSgroup = 0
	var thisErr = ""
	var totalchecked = 0
	stringToSplit = "<%
set groupRs=conn.execute("select uid from reason where reasonType=1")
if not groupRs.eof then
	response.write groupRs.getString(,,,",")
end if
groupRs.close
set groupRs=nothing
%>"
	groupOne = stringToSplit.split(",")
	for (var i = 0; i < formObj.TS.length; i++) {
		if (formObj.TS[i].checked){
			totalchecked += 1
			if (TSgroup==0){
				for (var j = 0; j < groupOne.length; j++){
					if (formObj.TS[i].value==groupOne[j])
						TSgroup=1
				}
				if (TSgroup==0)
					TSgroup=2
			}else{
				for (var j = 0; j < groupOne.length; j++){
					if (formObj.TS[i].value==groupOne[j]){
						if (TSgroup==2){
							thisErr="yes"
						}
					}
				}
			}
		}
	}
	if (thisErr=="yes"){
		reqField=reqField+", 同一種類的原因";
		placeFocus=formObj.otherReason1;
	}
	if (totalchecked==0&&formObj.otherReason1.value==''&&formObj.otherReason2.value==''){
		reqField=reqField+", 最少一個原因";
		placeFocus=formObj.otherReason1;
	}

	if (formObj.firstApproval.selectedIndex!=0&&formObj.firstApprovalDate.value==""){
		reqField=reqField+", 貸委會批核日期";
		if (!placeFocus)
			placeFocus=formObj.firstApprovalDate;
	}

	if (formObj.secondApproval.selectedIndex!=0&&formObj.secondApprovalDate.value==""){
		reqField=reqField+", 董事會批核日期";
		if (!placeFocus)
			placeFocus=formObj.secondApprovalDate;
	}

	if (formObj.loanPlanID.options[formObj.loanPlanID.selectedIndex].text=='聯名貸款'&&formObj.guarantorID.value==""){
		reqField=reqField+", 擔保人";
		if (!placeFocus)
			placeFocus=formObj.guarantorID;
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

function calculation(){
	formObj=document.form1;
	mInterest=0;loanAmt=0;installment=0
	if (formObj.loanAmt.value!=""&&formObj.loanAmt.value!=0&&formObj.installment.value!=""&&formObj.installment.value!=0&&formObj.chequeDate.value!=""){
		loanAmt=parseInt(formObj.loanAmt.value)
		installment=parseInt(formObj.installment.value)
		if (formObj.interest.value=="-1"){
			chequeDate=formObj.chequeDate.value
			Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
			Y=parseInt(chequeDate.substr(6,4))
			if (chequeDate.substr(3,1)=="0")
				M=parseInt(chequeDate.substr(4,1))
			else
				M=parseInt(chequeDate.substr(3,2))
			if (chequeDate.substr(0,1)=="0")
				D=parseInt(chequeDate.substr(1,1))
			else
				D=parseInt(chequeDate.substr(0,2))
			mD=Months[M-1]
			if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)) && M==2)
				mD=29
//			mInterest=Math.floor((((loanAmt+loanAmt/installment)/200+(loanAmt/100/mD*(mD-D)/installment))+0.04)*20)/20
			mInterest=Math.floor(Math.floor((((loanAmt+loanAmt/installment)/200+(loanAmt/100/mD*(mD-D)/installment))+0.04)*20)/20*100)/100
		}else{
			mInterest=0
		}
		document.all.tags( "td" )['totalInterest'].innerHTML=mInterest*installment;
		document.all.tags( "td" )['monthlyPrincipal'].innerHTML=loanAmt/installment;
		document.all.tags( "td" )['monthlyInterest'].innerHTML=mInterest;
		document.all.tags( "td" )['monthlyRepay'].innerHTML=Math.floor((mInterest+loanAmt/installment+9.99)/10)*10;
	}
	formObj=document.form1;
	if (formObj.netSalary.value!=""){
		ability=parseInt(formObj.netSalary.value)
		document.all.tags( "td" )['repayAbility'].innerHTML=ability/4;
		if (Math.floor((mInterest+loanAmt/installment+9.99)/10)*10<=ability/4){
			document.all.tags( "td" )['repayAbility'].style.color ='black'
		}else{
			document.all.tags( "td" )['repayAbility'].style.color ='red'
		}
	}
	if (formObj.guarantorSalary.value!=""){
		Gability=parseInt(formObj.guarantorSalary.value)
		document.all.tags( "td" )['GRepayAbility'].innerHTML=Gability/4;
		if (Math.floor((mInterest+loanAmt/installment+9.99)/10)*10<=Gability/4){
			document.all.tags( "td" )['GRepayAbility'].style.color ='black'
		}else{
			document.all.tags( "td" )['GRepayAbility'].style.color ='red'
		}
	}
}

function memberChange(){
	if (document.form1.memNo.value==''){
		document.form1.memName.value=''
		document.all.tags( "td" )['memName'].innerHTML=''
		document.form1.memGrade.value=''
		document.all.tags( "td" )['memGrade'].innerHTML=''
		document.form1.employCond.value=''
		document.all.tags( "td" )['employCond'].innerHTML=''
		document.form1.age.value=''
		document.all.tags( "td" )['age'].innerHTML=''
		document.form1.firstAppointDate.value=''
		document.all.tags( "td" )['firstAppointDate'].innerHTML=''
	}
	popup('pop_searchMem.asp?key='+document.form1.memNo.value)
}

function guarantorChange(){
	if (document.form1.guarantorID.value==''||document.form1.guarantorID.value=='0'){
		document.form1.guarantorName.value=''
		document.all.tags( "td" )['guarantorName'].innerHTML=''
		document.form1.guarantorGrade.value=''
		document.all.tags( "td" )['guarantorGrade'].innerHTML=''
	}else{
		popup('pop_searchGua.asp?key='+document.form1.guarantorID.value);
	}
}

function clearOthers(group){
	groupString = ",<%
set groupRs=conn.execute("select uid from reason where reasonType=1")
if not groupRs.eof then
	response.write groupRs.getString(,,,",")
end if
groupRs.close
set groupRs=nothing
%>"
	for (var i = 0; i < formObj.TS.length; i++) {
		var checkString = ","+formObj.TS[i].value+","
		if (group==1){
			if (groupString.indexOf(checkString)==-1)
				formObj.TS[i].checked=false;
		}
		if (group==2){
			if (groupString.indexOf(checkString)>=0)
				formObj.TS[i].checked=false;
		}
	}
	if (group==1){
		formObj.otherReason2.value='';
	}
	if (group==2){
		formObj.otherReason1.value='';
	}

}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="calculation();form1.appDate.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="loanDetail.asp">
<input type="hidden" name="uid" value="<%=uid%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memGrade" value="<%=memGrade%>">
<input type="hidden" name="age" value="<%=age%>">
<input type="hidden" name="employCond" value="<%=employCond%>">
<input type="hidden" name="firstAppointDate" value="<%=firstAppointDate%>">
<input type="hidden" name="guarantorName" value="<%=guarantorName%>">
<input type="hidden" name="guarantorGrade" value="<%=guarantorGrade%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="300" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="b8" align="right">日期</td>
					<td width=10></td>
					<td>
						<input type="text" name="appDate" value="<%=appDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};form1.firstApprovalDate.value=this.value">
						 (dd/mm/yy)
					</td>
				</tr>
				<tr>
					<td class="b8" align="right">申請類別</td>
					<td width=10></td>
					<td>
						<select name="loanType">
						<option value="N"<% if loanType="N" then %> selected<% end if%>>新申請
						<option value="E"<% if loanType="E" then %> selected<% end if%>>延期
						</select>
					</td>
				</tr>
				<tr>
					<td class="b8" align="right">社員編號</td>
					<td width=10></td>
					<td>
						<input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10" onchange="memberChange()">
						<input type="button" value="選擇" onclick="popup('pop_searchMem.asp?key='+document.form1.memNo.value)" class="sbttn">
					</td>
				</tr>
				<tr>
					<td class="b8" align="right">淨薪金</td>
					<td width=10></td>
					<td><input type="text" name="netSalary" value="<%=netSalary%>" size="17" maxlength="17" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
				</tr>
				<tr>
					<td class="b8" align="right">貸款額</td>
					<td width=10></td>
					<td><input type="text" name="loanAmt" value="<%=loanAmt%>" size="17" maxlength="17" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
				</tr>
				<tr>
					<td class="b8" align="right">攤分期數</td>
					<td width=10></td>
					<td><input type="text" name="installment" value="<%=installment%>" size="17" maxlength="17" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
				</tr>
				<tr>
					<td class="b8" align="right">預定支票日期</td>
					<td width=10></td>
					<td>
						<input type="text" name="chequeDate" value="<%=chequeDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};calculation();">
						 (dd/mm/yy)
					</td>
				</tr>
				<tr>
					<td class="b8" align="right">是否計息</td>
					<td width=10></td>
					<td>
						<select name="interest" onchange="calculation();">
						<option value="-1"<% if interest<>0 then %> selected<% end if%>>是</option>
						<option value="0"<% if interest=0 then %> selected<% end if%>>否</option>
						</select>
					</td>
				</tr>
			</table>
		</td>
		<td width="400" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
				<tr height="22">
					<td class="b8" align="right">姓名</td>
					<td width=10></td>
					<td id="memName"><%=memName%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">職級</td>
					<td width=10></td>
					<td id="memGrade"><%=memGrade%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">年齡</td>
					<td width=10></td>
					<td id="age"><%=age%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">招聘條款</td>
					<td width=10></td>
					<td id="employCond"><%=employCond%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">入職日期</td>
					<td width=10></td>
					<td id="firstAppointDate"><%=firstAppointDate%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">全期利息</td>
					<td width=10></td>
					<td id="totalInterest"></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">每月本金</td>
					<td width=10></td>
					<td id="monthlyPrincipal"></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">每月利息</td>
					<td width=10></td>
					<td id="monthlyInterest"></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">每月本利和</td>
					<td width=10></td>
					<td id="monthlyRepay"></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">還款能力</td>
					<td width=10></td>
					<td id="repayAbility"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<hr width="700">
<%
if uid="" then
	sql = "select b.reasonType,reasonName,a.uid,0 as theCheckPoint from reason a,reasonType b where a.reasonType=b.uid order by 1,2"
else
	sql = "select b.reasonType,a.reasonName,a.uid,1 as theCheckPoint from reason a,reasonType b, loanReason c "&_
	"where a.uid=c.reasonID and a.reasonType=b.uid and c.loanAppID="&uid&_
	"union select b.reasonType,reasonName,a.uid,0 from reason a,reasonType b "&_
	"where a.reasonType=b.uid and a.uid not in (select reasonID from loanReason where loanAppID="&uid&") order by 1,2"
end if
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn
thisType=0
%>
<table border="0" cellspacing="0" cellpadding="0">
<%
if not rs.eof then
%>
	<tr height="25" valign="bottom">
		<td colspan="4" class="b8"><%=rs("reasonType")%></td>
	</tr>
<%
	do while not rs.eof
		if rs("reasonType")<>"不時之需" then
			exit do
		end if
		if locate < 3 then locate=locate+1 else locate=0 end if
		if locate = 0 then
			response.write "<tr>"
		end if%>
		<td width="150" class="n8"><input type="checkbox" name="TS" value="<% =rs("uid") %>"<%if rs("theCheckPoint")<>0 then response.write " checked" end if%> onclick="if (this.checked){clearOthers(1)}"> <%=rs("reasonName")%>
		</td>
<%
		if locate = 3 then
			response.write "</tr>"
		end if
		rs.movenext
	loop
end if
%>
	<tr height="25" valign="bottom">
		<td class="b8">不時之需的其他原因</td>
		<td colspan="3"><input type="text" name="otherReason1" value="<%=otherReason1%>" size="50" maxlength="50" onchange="if(!this.value==''){clearOthers(1)}"></td>
	</tr>
<%
if not rs.eof then
	do while not rs.eof
		if thisType<>rs("reasonType") then
			thisType=rs("reasonType")
			locate = 3%>
			<tr height="25" valign="bottom">
				<td colspan="4" class="b8"><%=rs("reasonType")%></td>
			</tr>
	<%
		end if
		if locate < 3 then locate=locate+1 else locate=0 end if
		if locate = 0 then
			response.write "<tr>"
		end if%>
		<td width="150" class="n8"><input type="checkbox" name="TS" value="<% =rs("uid") %>"<%if rs("theCheckPoint")<>0 then response.write " checked" end if%> onclick="if (this.checked){clearOthers(2)}"> <%=rs("reasonName")%>
		</td>
<%
		if locate = 3 then
			response.write "</tr>"
		end if
		rs.movenext
	loop
end if
rs.close
%>
	<tr height="25" valign="bottom">
		<td class="b8">生產的其他原因</td>
		<td colspan="3"><input type="text" name="otherReason2" value="<%=otherReason2%>" size="50" maxlength="50" onchange="if(!this.value==''){clearOthers(2)}"></td>
	</tr>
</table>
<hr width="700">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">擔保人編號</td>
		<td width=10></td>
		<td>
			<input type="text" name="guarantorID" value="<%=guarantorID%>" size="10" onchange="guarantorChange()">
			<input type="button" value="選擇" onclick="popup('pop_searchGua.asp?key='+document.form1.guarantorID.value)" class="sbttn">
		</td>
		<td class="b8" align="right">擔保人姓名</td>
		<td width=10></td>
		<td width="200" id="guarantorName"><%=guarantorName%></td>
	</tr>
	<tr height="22">
		<td colspan="3"></td>
		<td class="b8" align="right">擔保人職級</td>
		<td width=10></td>
		<td id="guarantorGrade"><%=guarantorGrade%></td>
	</tr>
	<tr>
		<td class="b8" align="right">擔保人薪金</td>
		<td width=10></td>
		<td><input type="text" name="guarantorSalary" value="<%=guarantorSalary%>" size="17" maxlength="17" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
		<td class="b8" align="right">還款能力</td>
		<td width=10></td>
		<td id="GRepayAbility"></td>
	</tr>
	<tr>
		<td class="b8" align="right">會面日期</td>
		<td width=10></td>
		<td colspan="4">
			<input type="text" name="interviewDate" value="<%=interviewDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};">
			 (dd/mm/yy)
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">會面內容</td>
		<td width=10></td>
		<td colspan="4"><textarea rows="2" name="interviewDetail" cols="80"><%=interviewDetail%></textarea></td>
	</tr>
</table>
<hr width="700">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100" class="b8" align="right">貸委會批核</td>
		<td width=10></td>
		<td width="100">
			<select name="firstApproval">
			<option>
			<option<% if firstApproval="Approved" then %> selected<% end if%>>Approved
			<option<% if firstApproval="Rejected" then %> selected<% end if%>>Rejected
			<option<% if firstApproval="pending" then %> selected<% end if%>>pending
			</select>
		</td>
		<td width="40" class="b8" align="right">日期</td>
		<td width=10></td>
		<td width="200">
			<input type="text" name="firstApprovalDate" value="<%=firstApprovalDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};">
			 (dd/mm/yy)
		</td>
		<td width="100" class="b8" align="right">貸款計劃</td>
		<td width=10></td>
		<td>
			<select name="loanPlanID">
			<option value="0">
<%		SQL = "select * from loanPlan"
	        set lb = server.createobject("ADODB.Recordset")
	        lb.open sql, conn
				do while not lb.eof %>
			<option value="<%=lb("uid")%>"<% if lb("uid") = cdbl(loanPlanID) then %> Selected<% end if %>><% =lb("planName") %>
<%				lb.movenext
				loop
				Set lb = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">董事會批核</td>
		<td></td>
		<td>
			<select name="secondApproval">
			<option>
			<option<% if secondApproval="Approved" then %> selected<% end if%>>Approved
			<option<% if secondApproval="Rejected" then %> selected<% end if%>>Rejected
			</select>
		</td>
		<td class="b8" align="right">日期</td>
		<td></td>
		<td>
			<input type="text" name="secondApprovalDate" value="<%=secondApprovalDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};">
			 (dd/mm/yy)
		</td>
		<td width="100" class="b8" align="right">優惠計劃</td>
		<td width=10></td>
		<td>
			<select name="SpecialPlanID">
			<option value="0">
<%		SQL = "select * from specialPlan"
	        set lb = server.createobject("ADODB.Recordset")
	        lb.open sql, conn
				do while not lb.eof %>
			<option value="<%=lb("uid")%>"<% if lb("uid") = cdbl(SpecialPlanID) then %> Selected<% end if %>><% =lb("planName") %>
<%				lb.movenext
				loop
				Set lb = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">拒絕原因</td>
		<td></td>
		<td colspan="7">
			<input type="text" name="rejectReason" value="<%=rejectReason%>" size="50" maxlength="50">
		</td>
	</tr>
	<tr>
		<td colspan="9" align="right" valign="middle">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
<br>
</center>
</form>
</body>
</html>
