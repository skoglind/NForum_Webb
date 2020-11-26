<%
lMedlem = Request.QueryString("m")
Response.Redirect("/avdelning/medlem/minaspel.asp?m=" & lMedlem)
%>