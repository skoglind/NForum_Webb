<!--#INCLUDE FILE="../__INC/includes.asp"-->

<%

  If config_LockDown_Kommentarer Then Response.Redirect("../default.asp")
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lID       = GetQ("e","123",0)
  
    If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
      RS_Open 1, "SELECT * FROM cms_Kommentarer WHERE cID = " & CLng(lID), True
        If Not rsDB(1).EOF Then
          lReturnAvd = rsDB(1)("cAvdelning")
          lReturnID  = rsDB(1)("cBindID")
          rsDB(1).Delete
        End If
        
        If lReturnAvd > lstAvdelning(-1) Then Response.Redirect("../default.asp")
        avdData = Split(lstAvdelning(lReturnAvd), ";")
      RS_Close 1
    End If
  
  Response.Redirect("/avdelning/" & avdData(0) & "/" & avdData(1) & ".asp?e=" & CLng(lReturnID) & "#kommentarer")

%>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->