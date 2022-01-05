<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
rs.close

SQL = "select min(memNo) from memMaster where deleted=0"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
mFrom = rs(0)
SQL = "select max(memNo) from memMaster where deleted=0"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
mTo = rs(0)
rs.close

SQL = "select distinct memSection from memMaster order by memSection"
Set Rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
%>
<html>
<head>
<title>社員個人結算書</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
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

	if (formObj.mFrom.value==""||!formatNum(formObj.mFrom)){
		reqField=reqField+", 起始社員";
		if (!placeFocus)
			placeFocus=formObj.mFrom;
	}

	if (formObj.mTo.value==""||!formatNum(formObj.mTo)){
		reqField=reqField+", 終止社員";
		if (!placeFocus)
			placeFocus=formObj.mTo;
	}

	if (!formatDate(formObj.mDate)){
		reqField=reqField+", 結算書日期";
		if (!placeFocus)
			placeFocus=formObj.mDate;
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

function CA(){
    for (var i=0;i<document.form1.elements.length;i++){
        var e = document.form1.elements[i];
        if ((e.name != 'allbox') && (e.type=='checkbox')){
            e.checked = document.form1.allbox.checked;
        }
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mYear.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>社員個人結算書</b>
<form method="post" action="acBalPrint.asp" name="form1" onsubmit="return validating()">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" class="b8">年份</td>
		<td width="10"></td>
		<td>
			<select name="mYear" style="width:90px">
			<option value="this">今年
			<option value="last">去年
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">開始會計月</td>
		<td width="10"></td>
		<td>
			<select name="mStart" style="width:80px">
			<%
			for idx = 1 to 12
				response.write "<option>"&idx
			next
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">終止會計月</td>
		<td></td>
		<td>
			<select name="mEnd" style="width:80px">
			<%
			for idx = 1 to 12
				if idx=acPeriod then
					response.write "<option selected>"&idx
				else
					response.write "<option>"&idx
				end if
			next
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">起始社員</td>
		<td></td>
		<td>
			<input type="text" name="mFrom" value="<% =mFrom %>" maxlength=10 size=10>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">終止社員</td>
		<td></td>
		<td>
			<input type="text" name="mTo" value="<% =mTo %>" maxlength=10 size=10>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">結算書日期</td>
		<td></td>
		<td><input type="text" name="mDate" value="<%=mDate%>" size="10" maxlength="10"> (dd/mm/yyyy)</td>
	</tr>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td colspan="16" height="30" valign="top"><input name="allbox" type="checkbox" onClick="CA();"> <i>選擇全部</i> </td>
    </tr>
    <tr>
<%
locate=1
do while not rs.eof %>
        <td width="30"><input type="checkbox" name="TS" value="<% =rs("memSection") %>"></td>
        <td width="80"><%=rs("memSection")%></td>
<%
	if locate<8 then
		locate = locate + 1
	else
		locate = 1
		response.write "</tr><tr>"
	end if
	rs.movenext
loop%>
	</tr>
    <tr>
        <td colspan="16" height="30" valign="bottom" align="right"><input type="submit" value="確定" name="submit" class="sbttn" onclick="return validating()"></td>
    </tr>
</table>
</form>
</center></div>
</body>
</html>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
%>