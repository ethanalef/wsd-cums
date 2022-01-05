<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%


set rs = server.createobject("ADODB.Recordset")

sql = "select * from glControl"
rs.open sql, conn
acPeriod=rs("acPeriod")
rs.close



if request.form("action") = "Save" then

	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
         set rs = server.createobject("ADODB.Recordset")
        xdate =  year(sdate)&"/"&right("0"&month(sdate),2)&"/"&right("0"&day(sdate),2)
	sql = "select * from share where memno='"&memno&"' and date='"&xdate&"' and code ='"&code&"' "	
	rs.open sql, conn,1       
        if not rs.eof then        
	rs("amount") = samount
	rs.update
        end if
	addUserLog "add Saveing Withdraw Update"
	msg = "紀錄已更新"
          
end if


	
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNp.select();form1.memNo.focus();">
<!-- #include file="menu.asp" -->
<div><center><font size="3">退股修正</font></center></div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="acDetail.asp">
<table border="0" cellspacing="0" cellpadding="0">
<input type="hidden" name="uid" value="<%=uid%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="code" value="<%=code%>">
<input type="hidden" name="sdate" value="<%=sdate%>">
<input type="hidden" name="amount" value="<%=amount%>">

<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="300" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">

				<tr>
					<td class="b8" align="right">社員編號</td>
					<td width=10></td>
					<td>
						<input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10" onchange="memberChange()">
						<input type="button" value="選擇" onclick="popup('pop_srhMemSt.asp?key='+document.form1.memNo.value)" class="sbttn">
					</td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">姓名</td>
					<td width=10></td>
					<td id="memName"><%=memName%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">日期</td>
					<td width=10></td>
					<td id ="sdate" ><%=sdate%></td>
				</tr>
				<td class="b8" align="right">項目</td>
					<td width=10></td>
					<td  id ="code"><%=code%></td>
				</tr>
				</tr>
				<td class="b8" align="right">金額</td>
					<td width=10></td>
					<td  id ="amount"><%=amount%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">更改金額</td>
					<td width=10></td>
					<td><input type="text" name="samount" value="<%=amount%>" size="10" maxlength="10"></td>
				</tr>


		</td>
	</tr>
</table>
				<tr>
					<td colspan="3" align="right">
						<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
						<input type="submit" value="確定" onclick="return validating()&&confirm('Are you going to save the Record?')" name="action" class="sbttn">
						<%end if%>

						<input type="button" value="查詢" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value)" class="sbttn">
						
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
