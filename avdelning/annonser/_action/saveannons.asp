<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    Call start_Rec2Session("annons")
    
      lID       = GetF("e","123",0)
      
      sTitel    = GetF("aTitel","ABC",255)
      sTextM    = GetF("aTextM","ABC",20000)
      sTyp      = GetF("aTyp","123",0)
      sKategori = GetF("aKategori","123",0)
      sSold     = GetF("aSold","CHK",0)
      sSynlig   = GetF("aSynlig","CHK",0)
      
      If Len(Trim(sTitel)) < 1 Then Response.Redirect("../ny_annons.asp?e=" & CLng(lID) & "&fail=4")
      If Len(Trim(sTextM)) < 50 Then Response.Redirect("../ny_annons.asp?e=" & CLng(lID) & "&fail=1")
      If sTyp < 1 Or sTyp > lstKSTyp(-1) Then Response.Redirect("../ny_annons.asp?e=" & CLng(lID) & "&fail=3")
      If sKategori < 1 Or sKategori > lstKSKategori(-1) Then Response.Redirect("../ny_annons.asp?e=" & CLng(lID) & "&fail=2")
      
      If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
        RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksID = " & CLng(lID), True
      Else
        RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksSkapadAv = " & CLng(CONST_USERID) & " And ksID = " & CLng(lID), True
      End If
        If rsDB(1).EOF Then
          rsDB(1).AddNew
          rsDB(1)("ksSkapadDatum")  = Now
          rsDB(1)("ksSkapadAv")     = CLng(CONST_USERID)
        Else
          If rsDB(1)("ksSkapadDatum") + CLng(config_AdDays) =< Now Then
            rsDB(1)("ksSkapadDatum")  = Now
          End If
        End If
        
        rsDB(1)("ksTitel")     = sTitel
        rsDB(1)("ksTextM")     = sTextM
        rsDB(1)("ksTyp")       = CLng(sTyp)
        rsDB(1)("ksKategori1") = CLng(sKategori)
        
        If sSold Then
          rsDB(1)("ksStatus")  = 1
        Else
          rsDB(1)("ksStatus")  = 0
        End If
        
        If sSynlig Then
          rsDB(1)("ksSynlig")  = 1
        Else
          rsDB(1)("ksSynlig")  = 0
        End If
        
        rsDB(1).Update
      RS_Close 1
    
    Call stop_Rec2Session("annons")
    Call SayMe("Sparad","Din <strong>annons</strong> har nu sparats!", "/avdelning/annonser/minaannonser.asp")

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->