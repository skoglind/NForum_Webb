<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    Call start_Rec2Session("recension")
    
      lID       = GetF("e","123",0)

      sTextM    = GetF("rTextM","ABC",20000)
      lBetyg    = GetF("rBetyg","123",0)
      
      lCnt = Con.ExeCute("SELECT COUNT(tID) FROM cms_SpelTitlar WHERE tID = " & CLng(lID))(0)
      If lCnt <> 1 Then Response.Redirect("../spel.asp")
      
      If Len(Trim(sTextM)) < 50 Then Response.Redirect("../ny_recension.asp?e=" & CLng(lID) & "&fail=1")
      If lBetyg < 1 Or lBetyg > 10 Then Response.Redirect("../ny_recension.asp?e=" & CLng(lID) & "&fail=2")
    
      lSpelID = Con.ExeCute("SELECT tSpelID FROM cms_SpelTitlar WHERE tID = " & CLng(lID))(0)
      sTitel  = Con.ExeCute("SELECT tTitel FROM cms_SpelTitlar WHERE tID = " & CLng(lID))(0)
      lKonsol  = Con.ExeCute("SELECT sKonsol FROM cms_SpelTitlar LEFT JOIN cms_Spel ON sID = tSpelID WHERE tID = " & CLng(lID))(0)
    
      RS_Open 1, "SELECT * FROM cms_Recensioner WHERE 1 = 2", True
        
        rsDB(1).AddNew
        
          rsDB(1)("rTitel")       = sTitel
          rsDB(1)("rText")        = sTextM
          rsDB(1)("rSkapadAv")    = CONST_USERID
          rsDB(1)("rDatumSkapad") = Now
          rsDB(1)("rAnvandarRec") = True
          
          rsDB(1)("rBetyg")       = lBetyg
          rsDB(1)("rKategori")    = lKonsol
          
          'If CONST_PUBLISH Then
          '  rsDB(1)("rStatus")          = 4
          '  rsDB(1)("rPubliceradAv")    = CONST_USERID
          '  rsDB(1)("rDatumPublicerad") = Now
          'Else
            rsDB(1)("rStatus")      = 2
          'End If
          
          rsDB(1)("rSpelID")      = CLng(lSpelID)
        
        rsDB(1).Update
      
      RS_Close 1
    
    Call stop_Rec2Session("recension")
    Call SayMe("Sparad","Din <strong>Recension</strong> har nu sparats, den kommer nu kontrolleras av oss innan den publiceras!", "/avdelning/spel/nusparad.asp?e=" & CLng(lID))

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->