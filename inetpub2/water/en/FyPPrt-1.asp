<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%
      Nyear=year(date())
      paidday= dmy(date())
%>

<html>
<head>
<title>���~��(PDP)</title>
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
<br><b5>���~��(PDP)</b>
<form method="post" action="FyPdfprintn.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td><font size="3" >�p����</formt></td>
                    <td width="10"></td>
 		    <td>
		    <select name="accode">
                    <option>
		    <option<% if accode="9922" then %> selected<% end if%>>9999 - �u�@�H��(1-500)
                    <option<% if accode="9922" then %> selected<% end if%>>9999 - �u�@�H��(501-1000)
                    <option<% if accode="9933" then %> selected<% end if%>>9999 - �u�@�H��(1001-1500)
                    <option<% if accode="9944" then %> selected<% end if%>>9999 - �u�@�H��(1501-2000)
                    <option<% if accode="9955" then %> selected<% end if%>>9999 - �u�@�H��(2001-2500)
                    <option<% if accode="9966" then %> selected<% end if%>>9999 - �u�@�H��(2501-3000)  
					<option<% if accode="9977" then %> selected<% end if%>>9999 - �u�@�H��(3001-4000)  
                    <option<% if accode="9988" then %> selected<% end if%>>9999 - �u�@�H��(4001-5000)
					<option<% if accode="9999" then %> selected<% end if%>>9999 - �u�@�H��(5001-6000>)

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
		<td align="right" class="b12">�~��</td>
		<td></td>
		<td>

			<input type="text" value="<%=nyear%>" name="nyear" size="4"  >
		</td>
	</tr>
	<tr>
		<td align="right" class="b12">�Ѯ����</td>
		<td></td>
		 <td><input type="text" name="paidday" value="<%=paidday%>" size="10" maxlength="10" <%=working%> onblur="if(!formatDate(this)){this.value=''};callage();">(dd/mm/yyyy)</td>
	</tr>

	<tr>
		<td align="right" class="b12">�Ѯ��v</td>
		<td></td>
		<td>

			<input type="text" value="<%=nrate%>" name="nrate" size="4"  >
		</td>
	</tr>
	<tr>
		<td align="right" class="b12">��X</td>
		<td></td>
		<td>
			<select name="output" style="width:80px">
			<option value="html">Html
			
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="�T�w" name="submit" class="sbttn">
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