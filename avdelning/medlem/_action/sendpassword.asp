<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    sEpost  = Trim(LCase(GetF("epost","ABC",255)))
    
    If Len(sEpost) < 5              Then Response.Redirect("../glomtlosen.asp?fail=1")    ' En f�r kort e-postadress
    If InStr(1, sEpost, "@", 1) < 1 Then Response.Redirect("../glomtlosen.asp?fail=1")    ' Saknar @
    If InStr(1, sEpost, ".", 1) < 1 Then Response.Redirect("../glomtlosen.asp?fail=1")    ' Saknar .
    If MakeLegal(sEpost) <> sEpost  Then Response.Redirect("../glomtlosen.asp?fail=1")    ' Oglitiga tecken
    sEpost = MakeLegal(sEpost)
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aEpost = '" & sEpost & "'", True
    
      If Not rsDB(1).EOF Then
        ' Oki, s�nd nu ut ett mail
        
        sUserMail = rsDB(1)("aEpost")
        sUserName = rsDB(1)("aAnvNamn")
        
        sPassKey  = SlumpText(10, True)
        
        rsDB(1)("aPassKey")       = sPassKey
        rsDB(1)("aNewPass")       = True
        
        sHTML = sHTML & ""
        sHTML = sHTML & "<h3>Automatiskt utskick fr�n N-Forum.se, Gl�mt l�senordet?</h3>"
        sHTML = sHTML & "<p>Du har beg�rt att f� ett nytt l�senord till anv�ndaren som �r registrerad med denna e-postadress p� N-Forum.se, klicka p� l�nken nedan f�r att �ndra ditt l�senord.</p>"
        sHTML = sHTML & "<p><a href=""http://" & page_NForum & "/avdelning/medlem/nyttlosenord.asp?ua=" & sUserName & "&x=" & sPassKey & """>http://" & page_NForum & "/avdelning/medlem/nyttlosenord.asp?ua=" & sUserName & "&x=" & sPassKey & "</a></p>"
        sHTML = sHTML & "<p><strong>Anv�ndarnamn:</strong> " & sUserName & "</p>"
        sHTML = sHTML & "<p><strong>Nyckel:</strong> " & sPassKey & "</p>"
        sHTML = sHTML & "<p>Om du INTE har beg�rt att f� ett nytt l�senord, ignorera d� bara detta brev s� kommer ingen �ndring av ditt l�senord att ske.</p>"
        sHTML = sHTML & "<p><b>N-forum.se</b></p>"
        
        SendMyMail sHTML, "N-Forum.se - Automatiskt utskick, gl�mt l�senordet?", sUserMail
        
        rsDB(1).Update
      Else
        Response.Redirect("../glomtlosen.asp?fail=1")                                     ' Inga poster funna
      End If
    
    RS_Close 1
    
    Session.Value("form_saved") = True
    Call SayMe("Skickat","Ditt <strong>l�senord</strong> har nu skickats!", "/avdelning/medlem/glomtlosen.asp?yes=1")
  
  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->