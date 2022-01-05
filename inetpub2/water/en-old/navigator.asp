<%
function navigator(byval httphead)
        
	if instr(httphead,"?")>0 then qString="&" else qString="?" end if
	response.write "<table><tr><td align=center>"
	if cint(pageno) > 1 then
		response.write "<a href="""&httphead&qString&"page="&pageno-1&""">前一頁</a>&nbsp;"
	end if
        if cint(pagecount) > 10 then
           if cint(pageno)/10=int(cint(pageno)/10) then
 		xx = int(cint(pageno)/10)
           else
              xx = int(cint(pageno)/10)+1          
           end if


        
           for idx = (xx -1)*10+1 to xx*10           

					if idx = cint(pageno) then
						response.write "<font color=black style=""FONT: bold"">"&idx&"</font>&nbsp;"
					else
						response.write "<a href="""&httphead&qString&"page="&idx&""">"&idx&"</a>&nbsp;"
					end if
				next
        end if
	if cint(pageno)<pagecount then
		response.write "<a href="""&httphead&qString&"page="&pageno+1&""">後一頁</a>"
	end if
	response.write "</td></tr></table>"
end function
%>