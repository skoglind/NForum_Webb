<%@ Language=VBScript %>
<%
Response.Status="301 Moved Permanently"
e = Request.QueryString("e")
If Not IsNumeric(e) Then e = 0
Response.AddHeader "Location","http://www.n-forum.se/avdelning/konsol/konsol_visa_info.asp?e=" & e
%>