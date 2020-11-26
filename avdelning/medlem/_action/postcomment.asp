<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    lID       = GetF("e","123",0)
    lMedlem   = GetF("m","ABC",50)

    If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME

    If Not dbUserExists(lMedlem) Then Response.Redirect("/")
    anvID = GetIDFromUsername(lMedlem) 
    
    If config_LockDown_Feedback Then Response.Redirect("../default.asp?m=" & lMedlem)
    
    sTextM    = GetF("aMsg","ABC",1000)
    
    If Len(Trim(sTextM)) < 1 Then Response.Redirect("../omdome.asp?m=" & lMedlem)
    
    RS_Open 1, "SELECT * FROM cms_Feedback WHERE 1 = 2", True
      
      rsDB(1).AddNew
      
        rsDB(1)("fbTextM")         = sTextM
        rsDB(1)("fbAnv")           = CONST_USERID
        rsDB(1)("fbDatum")         = Now
        rsDB(1)("fbMedlem")        = CLng(anvID)
      
      rsDB(1).Update
    
    RS_Close 1

    Call SayMe("Sparad","Ditt <strong>domdöme</strong> har nu sparats!", "/avdelning/medlem/omdome.asp?m=" & lMedlem)

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->