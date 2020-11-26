<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    sEpost  = Trim(LCase(GetF("epost","ABC",255)))
    
    If Len(sEpost) < 5              Then Response.Redirect("../glomtlosen.asp?fail=1")    ' En för kort e-postadress
    If InStr(1, sEpost, "@", 1) < 1 Then Response.Redirect("../glomtlosen.asp?fail=1")    ' Saknar @
    If InStr(1, sEpost, ".", 1) < 1 Then Response.Redirect("../glomtlosen.asp?fail=1")    ' Saknar .
    If MakeLegal(sEpost) <> sEpost  Then Response.Redirect("../glomtlosen.asp?fail=1")    ' Oglitiga tecken
    sEpost = MakeLegal(sEpost)
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aEpost = '" & sEpost & "'", True
    
      If Not rsDB(1).EOF Then
        ' Oki, sänd nu ut ett mail
        
        sUserMail = rsDB(1)("aEpost")
        sUserName = rsDB(1)("aAnvNamn")
        
        sPassKey  = SlumpText(10, True)
        
        rsDB(1)("aPassKey")       = sPassKey
        rsDB(1)("aNewPass")       = True
        
        sHTML = sHTML & ""
        sHTML = sHTML & "<h3>Automatiskt utskick från N-Forum.se, Glömt lösenordet?</h3>"
        sHTML = sHTML & "<p>Du har begärt att få ett nytt lösenord till användaren som är registrerad med denna e-postadress på N-Forum.se, klicka på länken nedan för att ändra ditt lösenord.</p>"
        sHTML = sHTML & "<p><a href=""http://" & page_NForum & "/avdelning/medlem/nyttlosenord.asp?ua=" & sUserName & "&x=" & sPassKey & """>http://" & page_NForum & "/avdelning/medlem/nyttlosenord.asp?ua=" & sUserName & "&x=" & sPassKey & "</a></p>"
        sHTML = sHTML & "<p><strong>Användarnamn:</strong> " & sUserName & "</p>"
        sHTML = sHTML & "<p><strong>Nyckel:</strong> " & sPassKey & "</p>"
        sHTML = sHTML & "<p>Om du INTE har begärt att få ett nytt lösenord, ignorera då bara detta brev så kommer ingen ändring av ditt lösenord att ske.</p>"
        sHTML = sHTML & "<p><b>N-forum.se</b></p>"
        
        SendMyMail sHTML, "N-Forum.se - Automatiskt utskick, glömt lösenordet?", sUserMail
        
        rsDB(1).Update
      Else
        Response.Redirect("../glomtlosen.asp?fail=1")                                     ' Inga poster funna
      End If
    
    RS_Close 1
    
    Session.Value("form_saved") = True
    Call SayMe("Skickat","Ditt <strong>lösenord</strong> har nu skickats!", "/avdelning/medlem/glomtlosen.asp?yes=1")
  
  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->