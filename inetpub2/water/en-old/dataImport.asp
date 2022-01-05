<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request("process")<>"" then
	Server.ScriptTimeout = 3600
	set exconn = server.createobject("adodb.connection")
	exconn.open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & request("mdbFile")
	Set rstList = conn.OpenSchema(20)
	With rstList
	   Do While Not .EOF
	     If .Fields("TABLE_TYPE") = "TABLE" and .Fields("TABLE_NAME")<>"dtproperties" Then
			conn.execute("delete "&.Fields("TABLE_NAME"))
			set sourceRs = exconn.execute("select * from "&.Fields("TABLE_NAME"))
			Set targetRs = Server.CreateObject("ADODB.Recordset")
			sql = "select * from "&.Fields("TABLE_NAME")
			targetRs.open sql, conn, 2, 2
			do while not sourceRs.eof
				targetRs.addnew
				For Each Field in sourceRs.fields
					TheString = "targetRs(""" & Field.name & """) = sourceRs(""" & Field.name & """)"
					Execute(TheString)
				Next
				targetRs.update
				sourceRs.movenext
			loop
			targetRs.close
	     End If
	    .MoveNext
	   Loop
	End With
	addUserLog "Database Importing"

	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>資料庫輸入</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.mdbFile.value==""){
		reqField=reqField+", 資料檔";
		if (!placeFocus)
			placeFocus=formObj.mdbFile;
	}

	if (formObj.mdbFile.value.indexOf(".mdb")!=formObj.mdbFile.value.length-4){
		reqField=reqField+", access (mdb) 資料檔";
		if (!placeFocus)
			placeFocus=formObj.mdbFile;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>資料庫輸入</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>">
資料檔 <input TYPE="file" name="mdbFile">
<input type="submit" name="process" value="確定" onclick="return validating()&&confirm('確定輸入資料?')">
</form>
<p><font size=4 color=red>注意! 此程序會更換現存資料</font></p>
</center>
</body>
</html>
