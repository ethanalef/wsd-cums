<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "member.asp"
end if

id = request("id")

if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if id="" then
		sql = "select count(*) from memMaster where memNo=" & memNo
		rs.open sql, conn
		if rs(0) > 0 then
			msg = "�������X�w�g�s�b "
		end if
		rs.close
	end if

	if memGuarantorNo<>"" then
		sql = "select count(*) from memMaster where memNo=" & memGuarantorNo
		rs.open sql, conn
		if rs(0) = 0 then
			msg = msg & "�䤣���O�H�����s�� "
		end if
		rs.close
	end if

	if memGuarantor4Others<>"" and memGuarantor4Others<>"0" then
		sql = "select count(*) from memMaster where memNo=" & memGuarantor4Others
		rs.open sql, conn
		if rs(0) = 0 then
			msg = msg & "�䤣���O��L�������s�� "
		end if
		rs.close
	end if

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select * from memMaster where 0=1"
		else
			sql = "select * from memMaster where memNo=" & id
		end if
		rs.open sql, conn, 2, 2
		if id = "" then
			rs.addnew
			rs("memNo") = memNo
			rs("bankAccNo") = ""
			id = rs("memNo")
			addUserLog "Add Member"
		else
			addUserLog "Modify Member Detail"
		end if
		rs("memName") = memName
		rs("memAddr1") = memAddr1
		rs("memAddr2") = memAddr2
		rs("memAddr3") = memAddr3
		rs("memContactTel") = memContactTel
		rs("memMobile") = memMobile
		rs("memHKID") = memHKID
		rs("memGender") = memGender
		if memBday<>"" then rs("memBday") = right(memBday,4)&"/"&mid(memBday,4,2)&"/"&left(memBday,2) else rs("memBday")="" end if
		rs("memGrade") = memGrade
		rs("memSection") = memSection
		if memGuarantorNo<>"" then rs("memGuarantorNo") = memGuarantorNo else rs("memGuarantorNo")=0 end if
		if memGuarantor4Others<>"" then rs("memGuarantor4Others") = memGuarantor4Others else rs("memGuarantor4Others")=0 end if
		rs("personEntitled") = personEntitled
		rs("treasRefNo") = treasRefNo
		rs("employCond") = employCond
		if firstAppointDate<>"" then rs("firstAppointDate") = right(firstAppointDate,4)&"/"&mid(firstAppointDate,4,2)&"/"&left(firstAppointDate,2) else rs("firstAppointDate")="" end if
		if memDate<>"" then rs("memDate") = right(memDate,4)&"/"&mid(memDate,4,2)&"/"&left(memDate,2) else rs("memday")="" end if
		rs.update
		conn.committrans
		msg = "�����w��s"
	end if
else
	if id <> "" then
		sql = "select * from memMaster where memNo=" & id
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "member.asp"
		else
			For Each Field in rs.fields
				if Field.name="memBday" or Field.name="firstAppointDate" or Field.name="memDate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if
	else
		memDate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	end if
end if
%>
<html>
<head>
<title>������ƭץ�</title>
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
		reqField=reqField+", �����s��";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}

	if (formObj.memName.value==""){
		reqField=reqField+", �W��";
		if (!placeFocus)
			placeFocus=formObj.memName;
	}

	if (formObj.memHKID.value==""){
		reqField=reqField+", �����Ҹ��X";
		if (!placeFocus)
			placeFocus=formObj.memHKID;
	}

	if (!formatDate(formObj.memBday)){
		reqField=reqField+", �X�ͤ��";
		if (!placeFocus)
			placeFocus=formObj.memBday;
	}

	if (!formatDate(formObj.firstAppointDate)){
		reqField=reqField+", �J¾���";
		if (!placeFocus)
			placeFocus=formObj.firstAppointDate;
	}

	if (!formatDate(formObj.memDate)){
		reqField=reqField+", �J�����";
		if (!placeFocus)
			placeFocus=formObj.memDate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNo.focus()">
<!-- #include file="menu.asp" -->
<div align=right><a href="acDetail.asp?id=<%=request("id")%>">�ӤH��ץ�</a>&nbsp;&nbsp;</div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="memberDetail.asp">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">�������X</td>
		<td width=10></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="50"<%if id<>"" then response.write " onfocus=""form1.memName.focus();""" end if%>></td>
	</tr>
	<tr>
		<td class="b8" align="right">�W��</td>
		<td width=10></td>
		<td><input type="text" name="memName" value="<%=memName%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�a�}</td>
		<td width=10></td>
		<td><input type="text" name="memAddr1" value="<%=memAddr1%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td></td>
		<td width=10></td>
		<td><input type="text" name="memAddr2" value="<%=memAddr2%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td></td>
		<td width=10></td>
		<td><input type="text" name="memAddr3" value="<%=memAddr3%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�q�ܸ��X</td>
		<td width=10></td>
		<td><input type="text" name="memContactTel" value="<%=memContactTel%>" size="50" maxlength="20"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�ⴣ���X</td>
		<td width=10></td>
		<td><input type="text" name="memMobile" value="<%=memMobile%>" size="50" maxlength="20"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�����Ҹ��X</td>
		<td width=10></td>
		<td><input type="text" name="memHKID" value="<%=memHKID%>" size="50" maxlength="9"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�ʧO</td>
		<td width=10></td>
		<td>
			<select name="memGender">
			<option<%if memGender="M" then response.write " selected" end if%>>M</option>
			<option<%if memGender="F" then response.write " selected" end if%>>F</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">�X�ͤ�� (dd/mm/yy)</td>
		<td width=10></td>
		<td><input type="text" name="memBday" value="<%=memBday%>" size="50" maxlength="10"></td>
	</tr>
	<tr>
		<td class="b8" align="right">¾��</td>
		<td width=10></td>
		<td><input type="text" name="memGrade" value="<%=memGrade%>" size="50" maxlength="8"></td>
	</tr>
	<tr>
		<td class="b8" align="right">����</td>
		<td width=10></td>
		<td><input type="text" name="memSection" value="<%=memSection%>" size="50" maxlength="10"></td>
	</tr>
	<tr>
		<td class="b8" align="right">��O�H�����s��</td>
		<td width=10></td>
		<td><input type="text" name="memGuarantorNo" value="<%=memGuarantorNo%>" size="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">��O��L�������s��</td>
		<td width=10></td>
		<td><input type="text" name="memGuarantor4Others" value="<%=memGuarantor4Others%>" size="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�Q���w�H</td>
		<td width=10></td>
		<td><input type="text" name="personEntitled" value="<%=personEntitled%>" size="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�w�нs��</td>
		<td width=10></td>
		<td><input type="text" name="treasRefNo" value="<%=treasRefNo%>" size="50" maxlength="8"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�۸u����</td>
		<td width=10></td>
		<td><input type="text" name="employCond" value="<%=employCond%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�J¾��� (dd/mm/yy)</td>
		<td width=10></td>
		<td><input type="text" name="firstAppointDate" value="<%=firstAppointDate%>" size="50" maxlength="10"></td>
	</tr>
	<tr>
		<td class="b8" align="right">�J����� (dd/mm/yy)</td>
		<td width=10></td>
		<td><input type="text" name="memDate" value="<%=memDate%>" size="50" maxlength="10"></td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
			<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="��^" name="back" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
