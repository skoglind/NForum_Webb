<%
  sbTitel = Session.Value("trans_Titel")
  sbText  = Session.Value("trans_Text")
  sbLank  = Session.Value("trans_Lank")
  
  If Len(sbLank) < 1 Then Response.Redirect("/default.asp")
  
  Session.Value("trans_Titel") = ""
  Session.Value("trans_Text") = ""
  Session.Value("trans_Lank") = ""
  
  Response.Redirect(sbLank)
%>