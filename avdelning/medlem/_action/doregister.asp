<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    Call start_Rec2Session("reg")
  
    saNamn    = Trim(GetF("anvnamn","ABC",30))
    sEpost1   = Trim(LCase(GetF("epost1","ABC",255)))
    sEpost2   = Trim(LCase(GetF("epost2","ABC",255)))
    sSvar     = Trim(LCase(GetF("math","ABC",20)))
    bAvtal    = GetF("avtal","CHK",0)
    
    If config_LockDown_Registrering Then Response.Redirect("../registreradig.asp?fail=666")   ' Ingen registrering, HALT!!
    
    If Len(saNamn) < 1               Then Response.Redirect("../registreradig.asp?fail=1")    ' F�r kort anv�ndarnamn
    If MakeLegal(saNamn) <> saNamn   Then Response.Redirect("../registreradig.asp?fail=1")    ' Ogiltiga tecken
    If dbUserExists(saNamn)          Then Response.Redirect("../registreradig.asp?fail=2")    ' Anv�ndarnamnet upptaget
    saNamn = MakeLegal(saNamn)
    
    If Len(sEpost1) < 5              Then Response.Redirect("../registreradig.asp?fail=3")    ' En f�r kort e-postadress
    If MakeLegal(sEpost1) <> sEpost1 Then Response.Redirect("../registreradig.asp?fail=3")    ' Oglitiga tecken
    If MailIsValid(sEpost1)          Then Response.Redirect("../registreradig.asp?fail=3")    ' Oglitiga mail
    If sEpost1 <> sEpost2            Then Response.Redirect("../registreradig.asp?fail=5")    ' De st�mde inte
    If dbMailExists(sEpost1)         Then Response.Redirect("../registreradig.asp?fail=4")    ' E-postadressen upptagen
    sEpost1 = MakeLegal(sEpost1)
    
    If Text2Num(sSvar) <> CLng(Session.Value("svaret")) Then Response.Redirect("../registreradig.asp?fail=7") ' Fel svar p� fr�gan
    
    If Not bAvtal                    Then Response.Redirect("../registreradig.asp?fail=6")    ' Reglerna ej godk�nda
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE 1 = 2", True
    
        ' Allt godk�nt, registrerar nu anv�ndaren
        
        rsDB(1).AddNew
        
        sPassword     = SlumpText(8, False)
        sActivateKey  = SlumpText(15, True)
        
        sDBSalt1      = SlumpText(5, False)
        sDBSalt2      = SlumpText(5, False)
        
        rsDB(1)("aAnvNamn")         = saNamn
        rsDB(1)("aNamn")            = saNamn
        rsDB(1)("aEpost")           = sEpost1
        rsDB(1)("aSalt1")           = sDBSalt1
        rsDB(1)("aSalt2")           = sDBSalt2
        rsDB(1)("aTitelID")         = config_UserTitle
        rsDB(1)("aPassWd")          = MD5(config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPassword & "" & config_Hash_Salt_2 & "" & sDBSalt2)
        rsDB(1)("aNyttLosenord")    = True
        rsDB(1)("aMedlemSedan")     = Now
        rsDB(1)("aTimeStamp")       = Now
        rsDB(1)("aInloggadSenast")  = Now
        rsDB(1)("aAktiverad")       = False
        rsDB(1)("aNewActivation")   = True
        rsDB(1)("aNewDelivered")    = True
        rsDB(1)("aIn_IP_Reg")       = Left(Request.ServerVariables("REMOTE_ADDR"),20)
        rsDB(1)("aAktiveringskod")  = sActivateKey
        rsDB(1)("aBlockadTill")     = #2003-01-01 00:00:00#
        
        sHTML = sHTML & ""
        sHTML = sHTML & "<h3>Automatiskt utskick fr�n N-Forum.se, Aktivera din anv�ndare.</h3>"
        sHTML = sHTML & "<p>Tack f�r att du har valt att registrera dig som medlem p� N-Forum.se, klicka p� l�nken nedan f�r att bekr�fta denna e-postadress.</p>"
        sHTML = sHTML & "<p><a href=""http://" & page_NForum & "/avdelning/medlem/aktivera.asp?ua=" & saNamn & "&x=" & sActivateKey & """>http://" & page_NForum & "/avdelning/medlem/aktivera.asp?ua=" & saNamn & "&x=" & sActivateKey & "</a></p>"
        sHTML = sHTML & "<p><strong>Anv�ndarnamn:</strong> " & saNamn & "</p>"
        sHTML = sHTML & "<p><strong>L�senord:</strong> " & sPassword & "</p>"
        sHTML = sHTML & "<p>Om du INTE har har registrerat dig p� N-Forum.se, ignorera d� bara detta brev.</p>"
        sHTML = sHTML & "<p><b>N-forum.se</b></p>"
        
        SendMyMail sHTML, "N-Forum.se - Automatiskt utskick, Anv�ndaraktivering!", sEpost1
        
        rsDB(1).Update
        
        lUserID = rsDB(1)("aID")
    
    RS_Close 1
    
    ' ## Skicka PM d� anv�ndaren nu �r registrerad med all m�jlig info.
    RS_Open 1, "SELECT * FROM fsBB_PM WHERE 1 = 2", True
      rsDB(1).AddNew
      
        rsDB(1)("pTill")      = CLng(lUserID)
        rsDB(1)("pFran")      = CLng(config_WelcomePMFrom)
        rsDB(1)("pAmne")      = config_WelcomePMTitle
        rsDB(1)("pPM")        = config_WelcomePM
        rsDB(1)("pDatum")     = Now
        
      rsDB(1).Update
    RS_Close 1
    
    Call stop_Rec2Session("reg")
    Response.Redirect("../nuregistrerad.asp")
  
  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->