<%@ Language=VBScript %>
<%
Response.Status="301 Moved Permanently"
d = Request.QueryString("d")
Response.AddHeader "Location","http://www.n-forum.se/nintendo/kalender.asp?d=" & d
%>