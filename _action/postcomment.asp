<!--#INCLUDE FILE="../__INC/includes.asp"-->

  <%
    
    If config_LockDown_Kommentarer Then Response.Redirect("../default.asp")
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    lAvd      = GetF("avd","123",0)
    lID       = GetF("e","123",0)

    sTextM    = GetF("aMsg","ABC",1000)
    
    If lAvd > lstAvdelning(-1) Then Response.Redirect("../default.asp")
    avdData = Split(lstAvdelning(lAvd), ";")
    
    If Len(Trim(sTextM)) < 1 Then Response.Redirect("/avdelning/" & avdData(0) & "/" & avdData(1) & ".asp?e=" & CLng(lID) & "#kommentarer")
  
    RS_Open 1, "SELECT * FROM " & avdData(2) & " WHERE " & avdData(3) & " = " & CLng(lID), False
      If rsDB(1).EOF Then
        Response.Redirect("../default.asp")
      End If
    RS_Close 1
    
    RS_Open 1, "SELECT * FROM cms_Kommentarer WHERE 1 = 2", True
      
      rsDB(1).AddNew
      
        sTextM = TraceHyperlinks(sTextM)
        
        rsDB(1)("cTextM")         = sTextM
        rsDB(1)("cAnv")           = CONST_USERID
        rsDB(1)("cDatum")         = Now
        rsDB(1)("cAvdelning")     = CLng(lAvd)
        rsDB(1)("cBindID")        = CLng(lID)

      rsDB(1).Update
    
      newComId                  = rsDB(1)("cID")
      
    RS_Close 1

    Call SayMe("Sparad","Din <strong>kommentar</strong> har nu sparats!", "/avdelning/" & avdData(0) & "/" & avdData(1) & ".asp?e=" & CLng(lID) & "#kommentar_" & CLng(newComId))

  %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->