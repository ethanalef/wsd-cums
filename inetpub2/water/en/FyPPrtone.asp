<!-- #include file="../conn.asp" -->
<%
      Nyear=year(date())
%>

<html>
<head>
<title>全年結(Epson 890一人列表)</title>
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
<br><b5>全年結(PDP)</b>
<form method="post" action="Fyoneprtn.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td><font size="3" >聯絡員</formt></td>
                    <td width="10"></td>
 		    <td>
		    <select name="accode">
                    <option>
		    <option<% if accode="99977" then %> selected<% end if%>>9999 - 工作人員(1-15000)
      
                    <option<% if accode="9988" then %> selected<% end if%>>9999 - 工作人員(1501-3000)
                    <option<% if accode="9999" then %> selected<% end if%>>9999 - 工作人員(3001-5000)

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
		<td align="right" class="b12">社員編號</td>
		<td></td>
		<td>

			<input type="text" value="<%=memno%>" name="memno" size="4"  >
		</td>
	</tr>
	<tr>
		<td align="right" class="b12">年份</td>
		<td></td>
		<td>

			<input type="text" value="<%=nyear%>" name="nyear" size="4"  >
		</td>
	</tr>
	<tr>
		<td align="right" class="b12">股息率</td>
		<td></td>
		<td>

			<input type="text" value="<%=nrate%>" name="nrate" size="4"  >
		</td>
	</tr>
	<tr>
		<td align="right" class="b12">輸出</td>
		<td></td>
		<td>
			<select name="output" style="width:80px">
			<option value="html">Html
			
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