<!-- #include file="../CheckUserStatus.asp" -->
<html>
<head>
<title>�~���Ѯ��պ�</title>
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
		reqField=reqField+", �Ѯ��v";
		if (!placeFocus)
			placeFocus=formObj.mDividend;
	}else{
		if (!formatNum(formObj.mDividend)){
			reqField=reqField+", ���T�Ѯ��v";
			if (!placeFocus)
				placeFocus=formObj.mDividend;
		}
	}

    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "�ж�J"+reqField.substring(2);
        else
	        reqField = "�ж�J"+reqField.substring(2,reqField.lastIndexOf(","))+'��'+reqField.substring(reqField.lastIndexOf(",")+2);
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
<h3>�~���Ѯ��պ�</h3>
<form name="form1" method="post" action="yearEndReportPrint.asp" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">�Ѯ��v</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDividend" value="<%=mDividend%>" size="8" maxlength="8">%
		</td>
	</tr>
	<tr>
		<td class="b8">�~��</td>
		<td width="10"></td>
		<td>
			<select name="year" style="width:88px">
			<option value="this">���~
			<option value="last">�h�~
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8">��X</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="�T�w" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
