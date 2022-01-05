<%requiredLevel=3%>
<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select min(glId) as code from glMaster"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
mFrom = rs(0)
SQL = "select max(glId) as code from glMaster"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
mTo = rs(0)
rs.close
set rs=nothing

mStart="01/04/2003"
mEnd=right("0"&day(date),2)&"/"&right("0"&month(date),2)&"/"&year(date)
%>
<html>
<head>
<title>Transaction List</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
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
    if (sCount != 2){
      return false;
    }
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
    dateform.value = strD+'/'+strM+'/'+strY;
    if (!valDate(strM, strD, strY))
      return false;
    else
      return true;
  }
}

function validating1(){
	formObj=document.form1;
	place_focus=false;
	req_field="";

	if (formObj.mStart.value==""){
		if (req_field!="") req_field=req_field+", ";
		req_field=req_field+"start date";
		if (!place_focus){
			formObj.mStart.focus();
			place_focus=true;
		}
	}else{
		if (!formatDate(formObj.mStart)){
			if (req_field!="") req_field=req_field+", ";
			req_field=req_field+"correct start date format";
			if (!place_focus){
				formObj.mStart.focus();
				place_focus=true;
			}
		}
	}

	if (formObj.mEnd.value==""){
		if (req_field!="") req_field=req_field+", ";
		req_field=req_field+"end date";
		if (!place_focus){
			formObj.mEnd.focus();
			place_focus=true;
		}
	}else{
		if (!formatDate(formObj.mEnd)){
			if (req_field!="") req_field=req_field+", ";
			req_field=req_field+"correct end date format";
			if (!place_focus){
				formObj.mEnd.focus();
				place_focus=true;
			}
		}
	}

    if (req_field){
        req_field="Please fill in "+req_field+".";
        alert(req_field);
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" alink="#003399" link="#003399" vlink="#003399" bgcolor="#eeeef0" onload="form1.mStart.focus()">
<!-- #include file="menu.asp" -->
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<div align="center"><center>
<table border="0" cellpadding="0" cellspacing="0" width="900">
	<tr>
		<td colspan=2>
			<table border="0" width="100%">
				<tr>
					<td align=center>
						<br>
						<font color=#FF0000>*</font> field must be filled
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<form method="post" action="acListPrint.asp" name="form1" onsubmit="return validating1()">
	<input type=hidden name="from" value="<% =Request.servervariables("script_name") %>">
	<tr>
		<td valign="top" align="center">
			<table border="0" cellpadding="0" cellspacing="0" width="270">
				<tr>
					<td width="100" align="right" class="b8"><font color=#FF0000>* </font>Start Date</td>
					<td width="10"></td>
					<td width="160">
						<input type="text" name="mStart" value="<% =mStart %>" maxlength=10 size=10>
					</td>
				</tr>
				<tr>
					<td align="right" class="b8"><font color=#FF0000>* </font>End Date</td>
					<td></td>
					<td>
						<input type="text" name="mEnd" value="<% =mEnd %>" maxlength=10 size=10>
					</td>
				</tr>
				<tr>
					<td align="right" class="b8"><font color=#FF0000>* </font>Account From</td>
					<td></td>
					<td>
						<input type="text" name="mFrom" value="<% =mFrom %>" maxlength=8 size=10>
					</td>
				</tr>
				<tr>
					<td align="right" class="b8"><font color=#FF0000>* </font>Account To</td>
					<td></td>
					<td>
						<input type="text" name="mTo" value="<% =mTo %>" maxlength=8 size=10>
					</td>
				</tr>
				<tr>
					<td align="right" class="b8">Output</td>
					<td></td>
					<td>
						<select name="output" style="width:88px">
						<option value="html">Html
						<option value="text">Text
						<option value="word">Word
						<option value="excel">Excel
						</select>
						<input type="submit" value="Submit" name="submit" class="sbttn">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>