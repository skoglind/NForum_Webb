<%

  Response.Cookies("NFORUM")("A")  = ""
  Response.Cookies("NFORUM")("P")  = ""
  Response.Cookies("NFORUM").Domain = "n-forum.se" 
  Response.Cookies("NFORUM").Expires = DateAdd("d", -1, now)
  Session.Abandon
  Response.Redirect("/")

%>