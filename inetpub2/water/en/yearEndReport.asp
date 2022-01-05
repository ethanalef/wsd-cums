<!-- #include file="../CheckUserStatus.asp" -->
<html>
<head>
<title>年結股息試算</title>
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

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.mDividend.value==""){
		reqField=reqField+", 股息率";
		if (!placeFocus)
			placeFocus=formObj.mDividend;
	}else{
		if (!formatNum(formObj.mDividend)){
			reqField=reqField+", 正確股息率";
			if (!placeFocus)
				placeFocus=formObj.mDividend;
		}
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mDividend.focus()">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>年結股息試算</h3>
<form name="form1" method="post" action="yearEndReportPrint.asp" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">股息率</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDividend" value="<%=mDividend%>" size="8" maxlength="8">%
		</td>
	</tr>
	<tr>
		<td class="b8">年份</td>
		<td width="10"></td>
		<td>
			<select name="year" style="width:88px">
			<option value="this">今年
			<option value="last">去年
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8">輸出</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
