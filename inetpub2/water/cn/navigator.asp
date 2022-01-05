<%
function navigator(byval httphead)
	if instr(httphead,"?")>0 then qString="&" else qString="?" end if
	response.write "<table><tr><td align=center>"
	if cint(pageno) > 1 then
		response.write "<a href="""&httphead&qString&"page="&pageno-1&""">前一頁</a>&nbsp;"
	end if
	if pagecount > 40 then
		base = 20
		for idx = 1 to -int(-(pagecount/base))
			if idx = -int(-(cint(pageno)/base)) then
				if idx = -int(-(pagecount/base)) then maxpage=pagecount else maxpage=(idx*base) end if
				for ii = (idx*base)-(base-1) to maxpage
					if ii = cint(pageno) then
						response.write "<font color=black style=""FONT: bold"">"&ii&"</font>&nbsp;"
					else
						response.write "<a href="""&httphead&qString&"page="&ii&""">"&ii&"</a>&nbsp;"
					end if
				next
			else
				response.write "<a href="""&httphead&qString&"page="&(idx*base)-(base-1)&""">["&(idx*base)-(base-1)&"-"&(idx*base)&"]</a>&nbsp;"
			end if
		next
	elseif pagecount > 2 then
		for idx = 1 to pagecount
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