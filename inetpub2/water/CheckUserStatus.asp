<%
If session("username") = "" then
'	response.redirect "../timeout.asp"
end if

http_referer = lcase(Request.ServerVariables("HTTP_REFERER"))
http_host = lcase(Request.ServerVariables("HTTP_Host"))
If http_referer = "" then
	response.redirect "../illegal.asp"
elseif instr(1,http_referer,"http://" & http_host ) <> 1 then
	response.redirect "../illegal.asp"
end if



'Make sure this page is not cached
'Response.Expires = -1
'Response.ExpiresAbsolute = Now() - 2
'Response.AddHeader "pragma","no-cache"
'Response.AddHeader "cache-control","private"
'Response.CacheControl = "No-Store"

'if requiredLevel="" or session("userLevel") < requiredLevel then
'	response.redirect "../illegal.asp"
'end if
%>
