<%requiredLevel=3%>
<% 
  sdate1="01"&"/"&right("0"&month(date()),2)&"/"&(year(date()))
 sdate2=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
%>

<html>
<head>
<title>股金帳細項列表算</title>
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

	if (!formatDate(formObj.sdate2)){
		reqField=reqField+", 列印日期至";
		if (!placeFocus)
			placeFocus=formObj.sdate2;
	}

	if (!formatDate(formObj.sdate1)){
		reqField=reqField+", 開始列印日期由";
		if (!placeFocus)
			placeFocus=formObj.sdate1;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.sdate1.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>股金帳細項列表</b>
<form method="post" action="savtlstPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
        <tr>
		<td align="right" class="b8">開始列印日期由</td>
		<td width="10"></td>
		<td>
                <input type="text" name="sdate1" value="<%=sdate1%>" size="10" maxlenght="10" onblur="if(!formatDate(this)){this.value=''};">(dd/mm/yyyy)
                </td> 

        </tr>  
        <tr>
		<td align="right" class="b8">列印日期至</td>
		<td width="10"></td>
		<td>
                <input type="text" name="sdate2" value="<%=sdate2%>" size="10" maxlenght="10" onblur="if(!formatDate(this)){this.value=''};">(dd/mm/yyyy)
                </td> 

        </tr>
	<tr>
		<td align="right" class="b8">項目</td>
		<td width="10"></td>
		<td>
			<select name="KIND" style="width:88px">
			<option value="cash">現金
			<option value="bank">銀行
			<option value="Trea">庫房
			<option value="nacct">新戶
			<option value="swithd">退股
			<option value="ploan">退股還款
			<option value="Divid">股息
			<option value="cfee">會費			
			<option value="bfee">協會費
                        <option value="adj">調整
                        <option value="ins">保險金  
			<option value="all">全選
			</select>

		</td>
	</tr>
	<tr>
		<td align="right" class="b8">輸出</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定輸出?')"  class="sbttn">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>