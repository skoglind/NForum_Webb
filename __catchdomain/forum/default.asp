<%@ Language=VBScript %>
<%
i = Request.QueryString("tradid")
If Not IsNumeric(i) Then i = 0

Response.Status="301 Moved Permanently"
Response.AddHeader "Location","http://www.n-forum.se/avdelning/forum/trad.asp?e=" & CStr(i)
%> 