<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    Call start_Rec2Session("artikel")
    
      lID       = GetF("e","123",0)

      sTitel    = GetF("aTitel","ABC",255)
      sTextM    = GetF("aTextM","ABC",20000)
      lKonsol   = GetF("aKonsol","123",0)
      
      If Len(Trim(sTextM)) < 50 Then Response.Redirect("../ny_artikel.asp?fail=1")
      If lKonsol < 1 Or lKonsol > lstKonsol(0) Then Response.Redirect("../ny_artikel.asp?fail=2")
      If Len(Trim(sTitel)) < 5 Then Response.Redirect("../ny_artikel.asp?fail=3")
    
      RS_Open 1, "SELECT * FROM cms_Artiklar WHERE 1 = 2", True
        
        rsDB(1).AddNew
        
          rsDB(1)("aaTitel")       = sTitel
          rsDB(1)("aaText")        = sTextM
          rsDB(1)("aaSkapadAv")    = CONST_USERID
          rsDB(1)("aaDatumSkapad") = Now
          rsDB(1)("aaAnvandarArt") = True
          
          rsDB(1)("aaKategori")    = lKonsol
          
          'If CONST_PUBLISH Then
          '  rsDB(1)("aaStatus")          = 4
          '  rsDB(1)("aaPubliceradAv")    = CONST_USERID
          '  rsDB(1)("aaDatumPublicerad") = Now
          'Else
            rsDB(1)("aaStatus")      = 2
          'End If
        
        rsDB(1).Update
      
      RS_Close 1
    
    Call stop_Rec2Session("artikel")
    Call SayMe("Sparad","Din <strong>Artikel</strong> har nu sparats, den kommer nu kontrolleras av oss innan den publiceras!", "/avdelning/artiklar/nusparad.asp")

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->