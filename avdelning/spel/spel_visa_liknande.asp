<%@ Language=VBScript %>
<%
Response.Status="301 Moved Permanently"
e = Request.QueryString("e")
If Not IsNumeric(e) Then e = 0
Response.AddHeader "Location","http://www.n-forum.se/avdelning/spel/spel_visa_info.asp?e=" & e
%>