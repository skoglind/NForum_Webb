<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    sEpost  = Trim(LCase(GetF("epost","ABC",255)))
    
    If Len(sEpost) < 5              Then Response.Redirect("../omaktivera.asp?fail=1")    ' En för kort e-postadress
    If InStr(1, sEpost, "@", 1) < 1 Then Response.Redirect("../omaktivera.asp?fail=1")    ' Saknar @
    If InStr(1, sEpost, ".", 1) < 1 Then Response.Redirect("../omaktivera.asp?fail=1")    ' Saknar .
    If MakeLegal(sEpost) <> sEpost  Then Response.Redirect("../omaktivera.asp?fail=1")    ' Oglitiga tecken
    sEpost = MakeLegal(sEpost)
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aAktiverad = 0 AND aEpost = '" & sEpost & "'", True
    
      If Not rsDB(1).EOF Then
        ' Oki, sänd nu ut ett mail
        
        sUserMail = rsDB(1)("aEpost")
        sUserName = rsDB(1)("aAnvNamn")
        
        sPassword     = SlumpText(8, False)
        sActivateKey  = SlumpText(15, True)
        
        sDBSalt1      = SlumpText(5, False)
        sDBSalt2      = SlumpText(5, False)
        
        rsDB(1)("aSalt1")           = sDBSalt1
        rsDB(1)("aSalt2")           = sDBSalt2
        rsDB(1)("aPassWd")          = MD5(config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPassword & "" & config_Hash_Salt_2 & "" & sDBSalt2)
        rsDB(1)("aNyttLosenord")    = True
        rsDB(1)("aAktiveringskod")  = sActivateKey
        
        sHTML = sHTML & ""
        sHTML = sHTML & "<h3>Automatiskt utskick från N-Forum.se, Aktivera din användare.</h3>"
        sHTML = sHTML & "<p>Tack för att du har valt att registrera dig som medlem på N-Forum.se, klicka på länken nedan för att bekräfta denna e-postadress.</p>"
        sHTML = sHTML & "<p><a href=""http://" & page_NForum & "/avdelning/medlem/aktivera.asp?ua=" & sUserName & "&x=" & sActivateKey & """>http://" & page_NForum & "/avdelning/medlem/aktivera.asp?ua=" & sUserName & "&x=" & sActivateKey & "</a></p>"
        sHTML = sHTML & "<p><strong>Användarnamn:</strong> " & sUserName & "</p>"
        sHTML = sHTML & "<p><strong>Lösenord:</strong> " & sPassword & "</p>"
        sHTML = sHTML & "<p>Om du INTE har har registrerat dig på N-Forum.se, ignorera då bara detta brev.</p>"
        sHTML = sHTML & "<p><b>N-forum.se</b></p>"
        
        SendMyMail sHTML, "N-Forum.se - Automatiskt utskick, Användaraktivering!", sUserMail
        
        rsDB(1).Update
      Else
        Response.Redirect("../omaktivera.asp?fail=1")                                     ' Inga poster funna
      End If
    
    RS_Close 1
    
    Session.Value("form_saved") = True
    Call SayMe("Skickat","Ditt <strong>aktiveringsmail</strong> har nu skickats!", "/avdelning/medlem/omaktivera.asp?yes=1")
  
  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->