<%requiredLevel=3%>
<% 
 sdate1="01"&"/"&right("0"&month(date()),2)&"/"&year(date())
 sdate2=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
%>

<html>
<head>
<title>������b��ƦC��</title>
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

	if (!formatDate(formObj.sdate1)){
		reqField=reqField+", �}�l�C�L�����";
		if (!placeFocus)
			placeFocus=formObj.sdate1;
	}

	if (!formatDate(formObj.sdate2)){
		reqField=reqField+", �C�L�����";
		if (!placeFocus)
			placeFocus=formObj.sdate2;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.sdate1.focus()">
<!-- #include file="menu.asp" --> 
<div align="center"><center>
<br><b>������b��ƦC��</b>
<form method="post" action="MenDlstPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">

	<tr>
		<td align="right" class="b8">�����{�p</td>
		<td width="30"></td>
		<td>
			<select name="KIND" style="width:150px">
			<option value="N">���`
			<option value="L">�b�b
			<option value="D">�N��
			<option value="V">IVA
			<option value="C">�h��
 			<option value="x">�ᵲ
			<option value="p">�h�@
			<option value="B">�}��
			<option value="J">�s��
			<option value="T">�w��
			<option value="H">�Ȱ��Ȧ�
			<option value="A">�۰���b(ALL)
			<option value="0">�۰���b(�Ѫ�)
			<option value="1">�۰���b(�Ѫ�,�Q��)
			<option value"Z">�۰���b(�Ѫ�,����)
			<option value="3">�۰���b(�Q��,����)
			<option value="M">�w��,�Ȧ�
			<option value="F">�S�O�Ӯ�
                        <option value="8">�פ���y��b
                        <option value="9">�פ���y���`
			<option value="all">����
			</select>

		</td>
	</tr>


	<tr>
		<td align="right" class="b8">��X</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="�T�w" onclick="return validating()&&confirm('�T�w��X?')"  class="sbttn">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>