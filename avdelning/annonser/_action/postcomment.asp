<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If config_LockDown_Kommentarer Then Response.Redirect("../default.asp")
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    lID       = GetF("e","123",0)

    sTextM    = GetF("aMsg","ABC",1000)
    
    If Len(Trim(sTextM)) < 1 Then Response.Redirect("../annons_visa.asp?e=" & CLng(lID))
  
    RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksSkapadDatum + " & CLng(config_AdDays) & " > '" & Now & "' AND ksSynlig = 1 AND ksID = " & CLng(lID), False
      If rsDB(1).EOF Then
        Response.Redirect("../default.asp")
      Else
        lAnnonsAgare = rsDB(1)("ksSkapadAv")
      End If
    RS_Close 1
    
    RS_Open 1, "SELECT * FROM cms_Kommentar_KopSalj WHERE 1 = 2", True
      
      rsDB(1).AddNew
      
        rsDB(1)("kskTextM")         = sTextM
        rsDB(1)("kskAnv")           = CONST_USERID
        rsDB(1)("kskDatum")         = Now
        rsDB(1)("kskAnnons")        = CLng(lID)
      
      rsDB(1).Update
    
    RS_Close 1
    
    ' #### Skicka PM till annonsinnehavaren
    If CLng(lAnnonsAgare) <> CLng(CONST_USERID) Then SendPM CLng(lAnnonsAgare), CLng(CONST_USERID), "Automatiskt PM: Svar på annons", "[b]Automatiskt utskick av N-Forum.se[/b]" & vbCrlf & vbCrlf & "[I]Du har fått en kommentar på din annons av [B]" & getUserName(CONST_USERID) & "[/B] som följer nedan:[/I]" & vbCrlf & vbCrlf & sTextM & vbCrlf & vbCrlf & "» [url=/avdelning/annonser/annons_visa.asp?e=" & CLng(lID) & "]Visa annonsen[/url]"
    
    Call SayMe("Sparad","Din <strong>kommentar</strong> har nu sparats!", "/avdelning/annonser/annons_visa.asp?e=" & CLng(lID))

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->