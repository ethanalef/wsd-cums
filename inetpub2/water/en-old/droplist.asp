Sub Droplist(strSQL,strFieldName,strDefault,StrBoxName,strBoxTitle,strConn)

'Set Cursor
'-------------------------------------------------------------------------
Const adOpenStatic=3

' create the recordset, open it, and move to first record
'-------------------------------------------------------------------------

    Set rs = Server.CreateObject ("ADODB.Recordset")
    rs.Open strSQL, strConn,adOpenStatic
    
rs.movefirst

'Ouput result to droplist box
'-------------------------------------------------------------------------

strBoxTitle%>
<SELECT Name = <%=StrBoxName%> SIZE="1">
<OPTION SELECTED> <%=strDefault%> </OPTION>
<%do until rs.EOF%>
<OPTION> <%=rs(strFieldName)%> </OPTION>
<%rs.movenext
loop%>
</Select>

<%
'Close and clean up
'-------------------------------------------------------------------------
    rs.close
    set rs=nothing
End sub
%>
