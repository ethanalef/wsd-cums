<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->

<html>
<head>
<title>社員分組/組員列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function CA(){
    for (var i=0;i<document.form1.elements.length;i++){
        var e = document.form1.elements[i];
        if ((e.name != 'allbox') && (e.type=='checkbox')){
            e.checked = document.form1.allbox.checked;
        }
    }
}
//-->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b5>聯絡員列表</b>
<form method="post" action="sectionListPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td><font size="3" >聯絡員</formt></td>
                    <td width="10"></td>
 		    <td>
		    <select name="accode">
                    <option>
		    <option<% if accode="9999" then %> selected<% end if%>>9999 - 工作人員
<%
                     set rs=conn.execute("select  memno,memcname,memname,status from memmaster where  status='*'   order by memno  "    )
                         do while not rs.eof
                            if  rs(3)="*" then 
                            idx = rs(0)&"-"&rs(2)&" "&rs(1)
                      
%> 
		
			<option<% if accode=rs(0) then %> selected<% end if%>><%=idx%>
<%
                        end if               
                        rs.movenext
                        loop
                        rs.close 
			
%>                  
		    </select>
		    </td> 

		                             
               </tr>
	<tr>
		<td align="right" class="b12">輸出</td>
		<td></td>
		<td>
			<select name="output" style="width:80px">
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
</center></div>
</body>
</html>
<%

set rs = nothing
conn.close
set conn = nothing
%>