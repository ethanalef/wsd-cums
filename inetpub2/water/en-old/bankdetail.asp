<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->

<%


if request.form("back") <> "" then
	response.redirect "bank.asp"
end if
if request.form("new") <> "" then
	response.redirect "bankDetail.asp"
end if

uid  = request("bncode")


if request.form("bye") <>""  then
        id=""
 	For Each Field in Request.Form
		TheString = Field & "=id "
		Execute(TheString)
	Next

end if






if request.form("action") <> "" then

    
        
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
   	Next
        
        set rs = server.createobject("ADODB.Recordset")
	conn.begintrans
        if opt = 1 then
           conn.execute("update bank set bank = '"&bank&"' where bncode='"&uid&"' ")
        else
           conn.execute("insert into bank (bncode,bank) values ('"&bncode&"' ,'"&bank&"' )")
        end if



	conn.committrans
	msg = "紀錄已更新"

	set rs=nothing
        IF OPT = 1 THEN

	response.redirect "bank.asp"   
        END IF   
        uid=""
        bncode=""
        bank=""
else
	if uid <> "" then
                OPT = 1
		sql = "select * from bank where bncode="&uid
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "bank.asp"
		else
			For Each Field in rs.fields
				
					TheString = Field.name & "= rs(""" & Field.name & """)"
				
				Execute(TheString)
			Next
		end if
		rs.close

          
	else
                opt = 0
                bncode = ""
                bank =""

	end if
end if

%>
<html>
<head>
<title>銀行資料操作</title>
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


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.bncode.value==""){
		reqField=reqField+", 銀行號碼";
		if (!placeFocus)
			placeFocus=formObj.bncode;
	}

	if (formObj.bank.value==""){
		reqField=reqField+", 銀行名稱";
		if (!placeFocus)
			placeFocus=formObj.bank;
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


//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" <%if uid="" then response.write " onfocus=""form1.bncode.focus();""" else response.write " onfocus=""form1.bank.focus();""" end if%>>
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="bankDetail.asp">
<input type="hidden" name="OPT" value="<%=OPT%>">
<input type="hidden" name="uid" value="<%=uid%>">
<div><center><font size="3">銀行資料操</font></center></div>
<center>
<table border="0" cellspacing="0" cellpadding="0">
				<tr height="22">
					<td class="b8" align="right">銀行編號</td>
					<td width=10></td>
					<td><input type="text" name="bncode" value="<%=bncode%>" size="3" <%if uid<>"" then response.write " onfocus=""form1.bank.focus();""" end if%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">銀行名稱 </td>
					<td width=10></td>
					<td><input type="text" name="bank" value="<%=bank%>" size="50" ></td>
				</tr>
	<tr>
		<td colspan="9" align="right" valign="middle">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
<%if bncode="" then %>
		        <input type="submit" value="取消" name="bye"  class="sbttn">
<%end if %>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
<br>
</center>
</form>
</body>
</html>
