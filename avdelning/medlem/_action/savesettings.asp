<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    sP  = GetQ("p","ABC",100)
    lID = CONST_USERID 
    
    Call start_Rec2Session("settings")
    
    Select Case LCase(sP)
      Case "personlig"
        sNamn     = GetF("namn","ABC",50)
        sPlats    = GetF("plats","ABC",50)
        sHemsida  = GetF("hemsida","ABC",255)
        sMSN      = GetF("MSN","ABC",255)
        sICQ      = GetF("ICQ","ABC",50)
        sSignatur = GetF("signatur","ABC",50)
        
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), True
        
          rsDB(1)("aNamn")      = sNamn
          rsDB(1)("aPlats")     = sPlats
          rsDB(1)("aHemsida")   = sHemsida
          rsDB(1)("aMSN")       = sMSN
          rsDB(1)("aICQ")       = sICQ
          rsDB(1)("aSignatur")  = sSignatur
          
          rsDB(1).Update
        
        RS_Close 1
        
        Call stop_Rec2Session("settings")
        Session.Value("form_saved") = True
        Call SayMe("Sparad","Dina <strong>personliga inställningar</strong> har nu sparats!", "/avdelning/medlem/installningar.asp?p=personlig")
      Case "meddelande"
        sMsg      = GetF("PM","ABC",20000)
        
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), True
        
          rsDB(1)("aPM")      = sMsg
          
          rsDB(1).Update
        
        RS_Close 1
        
        Call stop_Rec2Session("settings")
        Session.Value("form_saved") = True
        Call SayMe("Sparad","Ditt <strong>personliga meddelande</strong> har nu sparats!", "/avdelning/medlem/installningar.asp?p=meddelande")
      Case "sidan"
        sPMPerSida      = GetF("PMSida","123",0)
        sPosition       = GetF("Position","123",0)
        sFont           = GetF("Font","123",0)
        sFontFam        = GetF("Fontfamily","123",0)
        bPM2Mail        = GetF("EpostPM","CHK",0)
        bAktiveraPM     = GetF("AktPM","CHK",0)
        lQuickList      = GetF("Quick","123",0)
        
        If sPMPerSida < 10 Or sPMPerSida > 50 Then sPMPerSida = 25
        If sPosition < 0 Or sPosition > 5     Then sPosition  = 0
        If sFont < 10 Or sFont > 14           Then sFont  = config_StandardSize
        If sFontFam < 1 Or sFontFam > 4       Then sFontFam  = config_StandardFont
        If lQuickList < 0 Or lQuickList > 2   Then lQuickList = 0
         
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), True
        
          rsDB(1)("aIn_PM")         = CLng(sPMPerSida)
          rsDB(1)("aIn_LoginPos")   = CLng(sPosition)
          rsDB(1)("aIn_Fontsize")   = CLng(sFont)
          rsDB(1)("aIn_Fontfamily") = CLng(sFontFam)
          rsDB(1)("aEpostPM")       = bPM2Mail
          rsDB(1)("aAktiveraPM")    = bAktiveraPM
          rsDB(1)("aIn_QuickList")  = lQuickList
          
          rsDB(1).Update
        
        RS_Close 1
        
        Session.Value("SET_PmSida")   = CLng(sPMPerSida)
        Session.Value("SET_FontSize") = CLng(sFont)
        Session.Value("SET_FontFam")  = CLng(sFontFam)
        Session.Value("NFORUM_PM")    = bAktiveraPM
        Session.Value("SET_Quick")    = CLng(lQuickList)
        
        Call stop_Rec2Session("settings")
        Session.Value("form_saved") = True
        Call SayMe("Sparad","Dina <strong>sidinställningar</strong> har nu sparats!", "/avdelning/medlem/installningar.asp?p=sidan")
      Case "forum"
        sTradarPerSida  = GetF("TradarSida","123",0)
        sInlaggPerSida  = GetF("InlaggSida","123",0)
        bVisaAvatar     = GetF("VisaAvatar","CHK",0)
        bVisaSign       = GetF("VisaSignatur","CHK",0)
        
        If sTradarPerSida < 10 Or sTradarPerSida > 50 Then sTradarPerSida = 25
        If sInlaggPerSida < 10 Or sInlaggPerSida > 40 Then sInlaggPerSida = 10
         
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), True
        
          rsDB(1)("aIn_Tradar")     = CLng(sTradarPerSida)
          rsDB(1)("aIn_Inlagg")     = CLng(sInlaggPerSida)
          rsDB(1)("aIn_Avatarer")   = bVisaAvatar
          rsDB(1)("aIn_Signaturer") = bVisaSign
          
          rsDB(1).Update
        
        RS_Close 1
        
        Session.Value("SET_TradarSida") = CLng(sTradarPerSida)
        Session.Value("SET_InlaggSida") = CLng(sInlaggPerSida)
        Session.Value("SET_ShowAvatar") = bVisaAvatar
        Session.Value("SET_ShowSign")   = bVisaSign
        
        Call stop_Rec2Session("settings")
        Session.Value("form_saved") = True
        Call SayMe("Sparad","Dina <strong>foruminställningar</strong> har nu sparats!", "/avdelning/medlem/installningar.asp?p=forum")
      Case "epost"
        sPass       = GetF("passwd","ABC",0)
        sEpost1     = GetF("epost1","ABC",255)
        sEpost2     = GetF("epost2","ABC",255)
        
        If Len(Trim(sEpost1)) < 5           Then Response.Redirect("../installningar.asp?p=epost&fail=2")    ' För kort
        If MakeLegal(sEpost1) <> sEpost1    Then Response.Redirect("../installningar.asp?p=epost&fail=2")    ' Ogiltiga tecken
        If MailIsValid(sEpost1)             Then Response.Redirect("../installningar.asp?p=epost&fail=2")    ' Spam-check failed
        If dbMailExists(sEpost1)            Then Response.Redirect("../installningar.asp?p=epost&fail=4")    ' Fanns redan
        If sEpost1 <> sEpost2               Then Response.Redirect("../installningar.asp?p=epost&fail=3")    ' De stämmer inte
        sEpost1 = Trim(MakeLegal(sEpost1))
         
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), True
        
          sDBSalt1  = rsDB(1)("aSalt1")
          sDBSalt2  = rsDB(1)("aSalt2")
          
          sMyPasswd    = rsDB(1)("aPassWd")
          sPassWd      = MD5(config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPass & "" & config_Hash_Salt_2 & "" & sDBSalt2)
          
          If LCase(sMyPasswd) = LCase(sPassWd) Then
            sActivateKey  = SlumpText(15, True)
          
            rsDB(1)("aNyEpost")        = sEpost1
            rsDB(1)("aAktiveradEpost") = sActivateKey
            
            saNamn = rsDB(1)("aAnvNamn")  
            
            sHTML = sHTML & ""
            sHTML = sHTML & "<h3>Automatiskt utskick från N-Forum.se, Byte av e-postadress.</h3>"
            sHTML = sHTML & "<p>Du har valt att ändra din e-postadress på N-forum.se till denna, klicka på länken nedan för att verifiera detta.</p>"
            sHTML = sHTML & "<p><a href=""http://" & page_NForum & "/avdelning/medlem/aktiveraepost.asp?ua=" & saNamn & "&x=" & sActivateKey & """>http://" & page_NForum & "/avdelning/medlem/aktiveraepost.asp?ua=" & saNamn & "&x=" & sActivateKey & "</a></p>"
            sHTML = sHTML & "<p>Om du INTE har bytt till denna e-postadress på N-forum.se ignorera då bara detta brev.</p>"
            sHTML = sHTML & "<p><b>N-forum.se</b></p>"
            
            SendMyMail sHTML, "N-Forum.se - Automatiskt utskick, Byte av e-postadress!", sEpost1
          
            rsDB(1).Update
          Else
            RS_Close 1
            Response.Redirect("../installningar.asp?p=epost&fail=1")   ' Den gamla är felaktigt
          End If
        
        RS_Close 1
        
        Call stop_Rec2Session("settings")
        Session.Value("form_saved") = True
        Call SayMe("Skickat","Din <strong>e-postadress</strong> kommer ändras när du klickat på länken i det utskickade verifieringsbrevet!", "/avdelning/medlem/installningar.asp?p=epost")
      Case "losenord"
        sOldPass    = GetF("oldpass","ABC",0)
        sPass1      = GetF("pass1","ABC",0)
        sPass2      = GetF("pass2","ABC",0)
        
        If Len(Trim(sPass1)) < 7  Then Response.Redirect("../installningar.asp?p=losenord&fail=2")   ' För kort
        If sPass1 <> sPass2       Then Response.Redirect("../installningar.asp?p=losenord&fail=3")   ' De stämmer inte
         
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), True
        
          sDBSalt1  = rsDB(1)("aSalt1")
          sDBSalt2  = rsDB(1)("aSalt2")
          sHash     = MD5(config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPass1 & "" & config_Hash_Salt_2 & "" & sDBSalt2)
          
          sMyOld    = rsDB(1)("aPassWd")
          sHashOld  = MD5(config_Hash_Salt_1 & "" & sDBSalt1 & "" & sOldPass & "" & config_Hash_Salt_2 & "" & sDBSalt2)
          
          If LCase(sMyOld) = LCase(sHashOld) Then
            rsDB(1)("aPassWd")        = sHash
            rsDB(1)("aNyttLosenord")  = True
          
            rsDB(1).Update
          Else
            RS_Close 1
            Response.Redirect("../installningar.asp?p=losenord&fail=1")   ' Det gamla är felaktigt
          End If
        
        RS_Close 1
        
        Call stop_Rec2Session("settings")
        Session.Value("form_saved") = True
        Call SayMe("Sparad","Ditt nya <strong>lösenord</strong> har nu sparats!", "/avdelning/medlem/installningar.asp?p=losenord")
      Case "deleteavatar"
        RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(CONST_USERID), True
          If rsDB(1)("aAvatar") Then
            Set fso = Server.CreateObject("Scripting.FileSystemObject")
              filename = Server.MapPath(config_Avatar) & "\u" & Right("000000" & CONST_USERID, 6) & ".jpg"
              If fso.FileExists(filename) Then fso.DeleteFile filename, True
            Set fso = Nothing
            
            rsDB(1)("aAvatar") = False
            rsDB(1).Update
          End If
        RS_Close 1
        
        Session.Value("form_saved") = True
        Call SayMe("Raderad","Din <strong>avatar</strong> har nu raderats!", "/avdelning/medlem/installningar.asp?p=avatar")
      Case Else
        Response.Redirect("../installningar.asp")
    End Select
  
  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->