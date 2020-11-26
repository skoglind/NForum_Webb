<%
lMedlem = Request.QueryString("m")
Response.Redirect("/avdelning/medlem/?m=" & lMedlem)
%>