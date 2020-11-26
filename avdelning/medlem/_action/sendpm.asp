<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    Call start_Rec2Session("pm")
      
      sTill       = UserIDFromName(GetF("pTill","ABC",100))
      sAmne       = GetF("pAmne","ABC",100)
      sTextM      = GetF("pMsg","ABC",20000)
      bAnswer     = GetF("pAnswer","CHK",0)
      
      If Not GetSendPM(sTill) Then Response.Redirect("../skrivpm.asp?fail=5")
      If Not CONST_PM         Then Response.Redirect("../skrivpm.asp?fail=6")
      
      If sTill = 0 Then Response.Redirect("../skrivpm.asp?fail=1")
      If sTill = CONST_USERID Then Response.Redirect("../skrivpm.asp?fail=2")
      If Len(Trim(sAmne)) < 1 Then Response.Redirect("../skrivpm.asp?fail=3")
      If Len(Trim(sTextM)) < 1 Then Response.Redirect("../skrivpm.asp?fail=4")
    
      SendPM CLng(sTill), CLng(CONST_USERID), sAmne, sTextM
      
    Call stop_Rec2Session("pm")
    Call SayMe("Skickat","Ditt <strong>PM</strong> har nu skickats!", "/avdelning/medlem/skickat.asp")

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->