<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%

server.scripttimeout = 1800

if request("process")<>"" then
	Select Case Request("output")
	Case "access"
		newFile = "export\backup"&year(date())&right("0"&month(date()),2)&right("0"&day(date()),2)&".mdb"
		thispath = Server.MapPath("../")
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.CopyFile thispath&"\blank\blank.mdb", thispath&"\"&newFile
		set exconn = server.createobject("adodb.connection")
		exconn.open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("..\"&newFile)
		Set rstList = conn.OpenSchema(20)
		With rstList
		   Do While Not .EOF
			 If .Fields("TABLE_TYPE") = "TABLE" and .Fields("TABLE_NAME")<>"dtproperties" Then
				set rs = conn.execute("select * from "&.Fields("TABLE_NAME"))
				Set exrs = Server.CreateObject("ADODB.Recordset")
				sql = "select * from "&.Fields("TABLE_NAME")
				exrs.open sql, exconn, 2, 2
				do while not rs.eof
					exrs.addnew
					For Each Field in rs.fields
						TheString = "exrs(""" & Field.name & """) = rs(""" & Field.name & """)"
						Execute(TheString)
					Next
					exrs.update
					rs.movenext
				loop
				exrs.close
			 End If
			.MoveNext
		   Loop
		End With
	Case "csv"
		newFile = "export\backup"&year(date())&right("0"&month(date()),2)&right("0"&day(date()),2)
		thispath = Server.MapPath("../")
		Set rstList = conn.OpenSchema(20)
		With rstList
		   Do While Not .EOF
			 If .Fields("TABLE_TYPE") = "TABLE" and .Fields("TABLE_NAME")<>"dtproperties" Then
				set rs = conn.execute("select * from "&.Fields("TABLE_NAME"))
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				Set objFile = objFSO.CreateTextFile(thispath&"\"&newfile&.Fields("TABLE_NAME")&".csv", True)
				For Each Field in rs.fields
					objFile.Write Field.name&","
				Next
				objFile.WriteLine ""
				do while not rs.eof
					For Each Field in rs.fields
						if instr(rs(Field.name),",")=0 then
							objFile.Write rs(Field.name)&","
						else
							objFile.Write """"&rs(Field.name)&""","
						end if
					Next
					objFile.WriteLine ""
					rs.movenext
				loop
				objFile.Close
			 End If
			.MoveNext
		   Loop
		End With
	Case "foxpro"
		thispath = Server.MapPath("../")
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.CopyFile thispath&"\blank\emsdcu.dbc", thispath&"\export\emsdcu.dbc"
		fso.CopyFile thispath&"\blank\emsdcu.dcx", thispath&"\export\emsdcu.dcx"
		fso.CopyFile thispath&"\blank\emsdcu.dct", thispath&"\export\emsdcu.dct"
		fso.CopyFile thispath&"\blank\cheque.DBF", thispath&"\export\cheque.DBF"
		fso.CopyFile thispath&"\blank\glControl.DBF", thispath&"\export\glControl.DBF"
		fso.CopyFile thispath&"\blank\glMaster.DBF", thispath&"\export\glMaster.DBF"
		fso.CopyFile thispath&"\blank\glTx.DBF", thispath&"\export\glTx.DBF"
		fso.CopyFile thispath&"\blank\handleParty.DBF", thispath&"\export\handleParty.DBF"
		fso.CopyFile thispath&"\blank\loanApp.DBF", thispath&"\export\loanApp.DBF"
		fso.CopyFile thispath&"\blank\loanApp.FPT", thispath&"\export\loanApp.FPT"
		fso.CopyFile thispath&"\blank\loanPlan.DBF", thispath&"\export\loanPlan.DBF"
		fso.CopyFile thispath&"\blank\loanReason.DBF", thispath&"\export\loanReason.DBF"
		fso.CopyFile thispath&"\blank\loginUser.DBF", thispath&"\export\loginUser.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes.DBF", thispath&"\export\meetingNotes.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes.FPT", thispath&"\export\meetingNotes.FPT"
		fso.CopyFile thispath&"\blank\meetingNotes0.DBF", thispath&"\export\meetingNotes0.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes1.DBF", thispath&"\export\meetingNotes1.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes2.DBF", thispath&"\export\meetingNotes2.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes3.DBF", thispath&"\export\meetingNotes3.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes4.DBF", thispath&"\export\meetingNotes4.DBF"
		fso.CopyFile thispath&"\blank\meetingNotes5.DBF", thispath&"\export\meetingNotes5.DBF"
		fso.CopyFile thispath&"\blank\memMaster.DBF", thispath&"\export\memMaster.DBF"
		fso.CopyFile thispath&"\blank\memTx.DBF", thispath&"\export\memTx.DBF"
		fso.CopyFile thispath&"\blank\reason.DBF", thispath&"\export\reason.DBF"
		fso.CopyFile thispath&"\blank\reasonType.DBF", thispath&"\export\reasonType.DBF"
		fso.CopyFile thispath&"\blank\specialPlan.DBF", thispath&"\export\specialPlan.DBF"
		fso.CopyFile thispath&"\blank\userLog.DBF", thispath&"\export\userLog.DBF"
		Set fso = fso.CreateTextFile(thispath&"\export\emsdcu-vfp.dsn", True)
		fso.WriteLine "[ODBC]"
		fso.WriteLine "DRIVER=Microsoft Visual FoxPro Driver"
		fso.WriteLine "UID="
		fso.WriteLine "Deleted=Yes"
		fso.WriteLine "Null=Yes"
		fso.WriteLine "Collate=Machine"
		fso.WriteLine "BackgroundFetch=Yes"
		fso.WriteLine "Exclusive=No"
		fso.WriteLine "SourceType=DBC"
		fso.WriteLine "SourceDB="&thispath&"\export\emsdcu.dbc"
		fso.Close
		set exconn = server.createobject("adodb.connection")
		exconn.open "filedsn="&thispath&"\export\emsdcu-vfp.dsn;DBQ="&thispath&"\export\emsdcu.dbc;UID=;PWD=;"
		Set rstList = conn.OpenSchema(20)
		With rstList
		   Do While Not .EOF
			 If .Fields("TABLE_TYPE") = "TABLE" and .Fields("TABLE_NAME")<>"dtproperties" Then
				set rs = conn.execute("select * from "&.Fields("TABLE_NAME"))
				Set exrs = Server.CreateObject("ADODB.Recordset")
				sql = "select * from "&.Fields("TABLE_NAME")
				exrs.open sql, exconn, 2, 2
				do while not rs.eof
					exrs.addnew
					For Each Field in rs.fields
						TheString = "exrs(""" & Field.name & """) = rs(""" & Field.name & """)"
						Execute(TheString)
					Next
					exrs.update
					rs.movenext
				loop
				exrs.close
			 End If
			.MoveNext
		   Loop
		End With
	Case "excel"
		newFile = "backup"&year(date())&right("0"&month(date()),2)&right("0"&day(date()),2)&".xls"
		thispath = Server.MapPath("../")
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.CopyFile thispath&"\blank\blank.xls", thispath&"\export\"&newFile
		set exconn = server.createobject("adodb.connection")
		exconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("..\export\"&newFile) & ";Extended Properties=""Excel 8.0;"""
		Set rstList = conn.OpenSchema(20)
		With rstList
		   Do While Not .EOF
			 If .Fields("TABLE_TYPE") = "TABLE" and .Fields("TABLE_NAME")<>"dtproperties" Then
				set rs = conn.execute("select * from "&.Fields("TABLE_NAME"))
				Set exrs = Server.CreateObject("ADODB.Recordset")
				sql = "select * from ["&.Fields("TABLE_NAME") & "$]"
				exrs.open sql, exconn, 2, 2
				do while not rs.eof
					exrs.addnew
					For Each Field in rs.fields
						TheString = "exrs(""" & Field.name & """) = rs(""" & Field.name & """)"
						Execute(TheString)
					Next
					exrs.update
					rs.movenext
				loop
				exrs.close
			 End If
			.MoveNext
		   Loop
		End With
	End Select
	addUserLog "Database Exporting"
	response.redirect "completed.asp"
end if
%>
<html>
<head>
<title>輸出資料</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<h3>輸出資料</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>">
	<tr>
		<td align="right" class="b8">輸出</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="access">Access
			<option value="csv">CSV
			<option value="foxpro">Foxpro
			<option value="excel">Excel
			</select>
			<input type="submit" value="確定" name="process" class="sbttn">
		</td>
	</tr>
</form>
</center>
</body>
</html>
