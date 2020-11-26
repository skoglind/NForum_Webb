<%
  Dim Con, Con_Status, rsDB(11), sessRec
  
  Function Con_Open()
    If Not Con_Status Then
      Con_Status = True
      
      Set con = Server.CreateObject("ADODB.Connection")
      con.Mode = 3
      con.Open config_ConnectionString
    End If
  End Function
  
  Function Con_Close()
    If Con_Status Then
      Con_Status = False
      
      Con.Close
      Set Con = Nothing
    End If
  End FUnction
  
  Function RS_Open(id, SQL, bEditable)
    Set rsDB(id) = Server.CreateObject("ADODB.RecordSet")
    If bEditable Then
      rsDB(id).Open SQL, Con, 1, 3
    Else
      rsDB(id).Open SQL, Con, 0, 1
    End If
  End Function
  
  Function RS_Close(id)
    rsDB(id).Close
    Set rsDB(id) = Nothing
  End Function
  
  Function LoginUser(sUserName, sPassword, bCookie, bSetCookie, sPostback)
    lFAIL     = 0
    
    If Len(Trim(sPostback)) < 1 Then sPostBack = "/"
    Session.Value("login_PB") = sPostBack
    
    If Len(sUserName) < 0 Or Len(sPassword) < 0 Then  lFAIL = 1  ' Lösenord och/eller användarnamn ej angivet
    If sUserName <> MakeLegal(sUserName) Then         lFAIL = 1  ' Ogiltigt användarnamn
    
    If lFAIL = 0 Then
      sUserName  = MakeLegal(sUserName)
      RS_Open 1, "SELECT * FROM fsBB_Anv LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID WHERE aAnvNamn = '" & sUserName & "'", True
    
      If rsDB(1).EOF Then                             lFAIL = 1  ' Användaren finns inte
      If lFAIL = 0 Then
        If rsDB(1)("aBlockadTill") > Now Then           lFAIL = 2  ' Användaren är bannad
        If Not rsDB(1)("aAktiverad") Then               lFAIL = 3  ' Användaren är inte aktiverad
        
        If lFAIL = 0 Then
          If bCookie Then
            ' USE COOKIE
            If Len(rsDB(1)("aPassWd")) < 10    Then lFAIL = 1  ' Felaktigt lösenord
            If rsDB(1)("aPassWd") <> sPassword Then lFAIL = 1  ' Felaktigt lösenord
          Else
            If rsDB(1)("aNyttLosenord") Then
              sDBSalt1  = rsDB(1)("aSalt1")
              sDBSalt2  = rsDB(1)("aSalt2")
              sHash     = config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPassword & "" & config_Hash_Salt_2 & "" & sDBSalt2
              sHash     = MD5(sHash)
            Else
              sHash = MD5(sPassword)
            End If
            
            If rsDB(1)("aPassWd") <> sHash Then            lFAIL = 1  ' Felaktigt lösenord
          End If
          
          If lFAIL = 0 Then 
            ' #### INLOGGNINGEN LYCKADES
            If Not rsDB(1)("aNyttLosenord") Then
              rsDB(1)("aNyttLosenord")  = True
              
              nSalt1                    = SlumpText(5, False)
              nSalt2                    = SlumpText(5, False)
              
              rsDB(1)("aSalt1")         = nSalt1
              rsDB(1)("aSalt2")         = nSalt2
              rsDB(1)("aPassWd")        = MD5(config_Hash_Salt_1 & "" & nSalt1 & "" & sPassword & "" & config_Hash_Salt_2 & "" & nSalt2)
            End If
            
              rsDB(1)("aInloggadSenast")  = Now
              rsDB(1)("aTimeStamp")       = Now
            rsDB(1).Update
            
            ' #### SÄTTER VAWRIABLER
            Session.Value("NFORUM_Login")       = True
            Session.Value("NFORUM_ID")          = rsDB(1)("aID")
            Session.Value("NFORUM_AnvNamn")     = rsDB(1)("aAnvNamn")
            Session.Value("NFORUM_TitelID")     = rsDB(1)("ttID")
            Session.Value("NFORUM_Admin")       = rsDB(1)("ttAdmin")
            Session.Value("NFORUM_DaysMember")  = DateDiff("d",rsDB(1)("aMedlemSedan"), Now)
            
            Session.Value("NFORUM_CMS")         = rsDB(1)("aS_CMS")
            Session.Value("NFORUM_CMS_RIGHTS")  = rsDB(1)("aS_CMSRatter")
            
            Session.Value("NFORUM_Publish")     = rsDB(1)("aDirektPublish")
            Session.Value("NFORUM_PM")          = rsDB(1)("aAktiveraPM")
            
            Session.Value("SET_TradarSida")     = rsDB(1)("aIn_Tradar")
            Session.Value("SET_InlaggSida")     = rsDB(1)("aIn_Inlagg")
            Session.Value("SET_PmSida")         = rsDB(1)("aIn_PM")
            
            Session.Value("SET_FontSize")       = rsDB(1)("aIn_Fontsize")
            Session.Value("SET_FontFam")        = rsDB(1)("aIn_Fontfamily")
            
            Session.Value("SET_AllowTimer")     = rsDB(1)("aHaveTimer")
            
            Session.Value("SET_ShowAvatar")     = rsDB(1)("aIn_Avatarer")
            Session.Value("SET_ShowSign")       = rsDB(1)("aIn_Signaturer")
            
            Session.Value("SET_Quick")          = rsDB(1)("aIn_QuickList")
            
            If bSetCookie Then
              ' MAKE COOKIE
              Response.Cookies("NFORUM")("A")  = sUserName
              Response.Cookies("NFORUM")("P")  = rsDB(1)("aPassWd")
              Response.Cookies("NFORUM").Domain = "n-forum.se" 
              Response.Cookies("NFORUM").Expires = DateAdd("d", Now, 365)
            End IF
          End If
        End If
      End If
      
      RS_Close 1
    End If
    
    If lFAIL > 0 Then
      If bCookie Then
        ' DELETE COOKIE!!
        Response.Cookies("NFORUM")("A")  = ""
        Response.Cookies("NFORUM")("P")  = ""
        Response.Cookies("NFORUM").Domain = "n-forum.se" 
        Response.Cookies("NFORUM").Expires = DateAdd("d", -1, now)
      Else
        Response.Redirect("/avdelning/medlem/loggain.asp?fail=" & lFAIL)  
      End if
    Else
      Session.Value("login_PB") = ""
      If Not bCookie Then
        Response.Redirect(sPostback)
      End If
    End If
  End Function
  
  Function start_Rec2Session(sStarter)
    If Len(Trim(sStarter)) > 0 Then 
      Session.Value("record_" & sStarter) = True
      sessRec = Trim(sStarter)
    End If
  End Function
  
  Function stop_Rec2Session(sStarter)
    sStarter = Trim(sStarter)
  
    If Len(sStarter) > 0 Then

      For Each Sess In Session.Contents
        If Left(LCase(Sess), Len(sStarter)) = sStarter Then Session.Contents.Remove(Sess)
      Next
      
      Session.Value("record_" & sStarter) = False
      sessRec = ""
    End If
  End Function
  
  Function GetQ(sName, sType, lLength)
    nVar = Request.QueryString(sName)
  
    Select Case Trim(UCase(sType))
      Case "ABC"
        nVar = Trim(nVar & " ")
        If lLength > 0 Then If Len(nVar) > lLength Then nVar = Left(nVar, lLength)
      Case "123"
        If Not IsNumeric(nVar) Or nVar = Empty Then nVar = 0
        nVar = CLng(nVar)
      Case "CHK"
        If Trim(UCase(nVar) & " ") = "YES" Then
          nVar = True
        Else
          nVar = False
        End If
    End Select
    
    If Len(sessRec) > 0 Then Session.Value(sessRec & "_" & sName) = nVar
    
    GetQ = nVar
  End Function
  
  Function GetF(sName, sType, lLength)
    nVar = Request.Form(sName)
  
    Select Case Trim(UCase(sType))
      Case "ABC"
        nVar = Trim(nVar & " ")
        If lLength > 0 Then If Len(nVar) > lLength Then nVar = Left(nVar, lLength)
      Case "123"
        If Not IsNumeric(nVar) Or nVar = Empty Then nVar = 0
        nVar = CLng(nVar)
      Case "CHK"
        If Trim(UCase(nVar) & " ") = "YES" Then
          nVar = True
        Else
          nVar = False
        End If
    End Select
    
    If Len(sessRec) > 0 Then Session.Value(sessRec & "_" & sName) = nVar
    
    GetF = nVar
  End Function
  
  Function SendPM(lTill, lFran, sAmne, sTextM)
    RS_Open 1, "SELECT * FROM fsBB_PM WHERE 1=2", True

      rsDB(1).AddNew
      
        rsDB(1)("pTill")      = CLng(lTill)
        rsDB(1)("pFran")      = CLng(lFran)
        rsDB(1)("pAmne")      = sAmne
        rsDB(1)("pPM")        = sTextM
        rsDB(1)("pDatum")     = Now
    
      rsDB(1).Update
    
      sPMID = rsDB(1)("pID")
    
    RS_Close 1
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lTill), False

      If Not rsDB(1).EOF Then
        If rsDB(1)("aEpostPM") Then
          sHTML = sHTML & ""
          sHTML = sHTML & "<h3>Automatiskt utskick från N-Forum.se, Du har fått ett PM.</h3>"
          sHTML = sHTML & "<p>Du har fått ett PM till din registrerade användare på N-Forum.se, klicka på länken nedan för att läsa ditt PM:</p>"
          sHTML = sHTML & "<p><a href=""http://" & page_NForum & "/avdelning/medlem/pm_visa.asp?e=" & sPMID & """>http://" & page_NForum & "/avdelning/medlem/pm_visa.asp?e=" & sPMID & "</a></p>"
          sHTML = sHTML & "<p>Om du inte vill bli meddelad per e-post vid nya PM kan du stänga av detta under dina inställningar på <a href=""http://" & page_NForum & "/"">N-Forum.se</a></p>"
          sHTML = sHTML & "<p><b>N-forum.se</b></p>"
        
          SendMyMail sHTML, "N-Forum.se - Automatiskt utskick, Du har fått ett PM!", rsDB(1)("aEpost")
        End If
      End If
      
    RS_Close 1
  End Function
  
  Function getRegKod(regKod)
    regKod = UCase(MakeLegal(regKod))
    bSure = 0
    
    Set ra = New RegExp 
      ra.Global = True
      ra.IgnoreCase = True
      
      ra.Pattern = "([ABCDEFGHIJKLMNOPQRSTUVWXYZ]{0,4})[-\s\.]{0,1}([ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789]{0,4})[-\s\.]{0,1}([ABCDEFGHIJKLMNOPQRSTUVWXYZ]{0,3})(.*)"
      
      reConsole = ra.Replace(regKod, "$1")
      reGame = ra.Replace(regKod, "$2")
      reRegion = ra.Replace(regKod, "$3")
    Set ra = Nothing
    
    If Len(reConsole) > 0 And Len(reGame) > 0 Then
      qNoConsole = "-" & reGame
      qSimple = reConsole & "-" & reGame
      If Len(reRegion) > 0 Then qFull = qSimple & "-" & reRegion
      
      If Len(qFull) > 0 Then
        lCnt_Correct = Con.ExeCute("SELECT COUNT(tID) FROM cms_SpelTitlar WHERE tRegionsKod = '" & qFull & "'")(0)
        lCnt_Almost = Con.ExeCute("SELECT COUNT(tID) FROM cms_SpelTitlar WHERE tRegionsKod LIKE '" & qFull & "%'")(0)
      End If
      
      lCnt_MayBe = Con.ExeCute("SELECT COUNT(tID) FROM cms_SpelTitlar WHERE tRegionsKod LIKE '%" & qNoConsole & "%'")(0)
      
      If lCnt_Correct > 0 Then
        SQL = "SELECT * FROM cms_SpelTitlar LEFT JOIN cms_Spel ON sID = tSpelID WHERE tRegionsKod = '" & qFull & "'"
        bSure = 1
      ElseIf lCnt_Almost > 0 Then
        SQL = "SELECT * FROM cms_SpelTitlar LEFT JOIN cms_Spel ON sID = tSpelID WHERE tRegionsKod LIKE '" & qFull & "%'"
      ElseIf lCnt_MayBe > 0 Then
        SQL = "SELECT * FROM cms_SpelTitlar LEFT JOIN cms_Spel ON sID = tSpelID WHERE tRegionsKod LIKE '%" & qNoConsole & "%'"
      Else
        SQL = "SELECT * FROM cms_SpelTitlar LEFT JOIN cms_Spel ON sID = tSpelID WHERE tRegionsKod LIKE '" & qSimple & "%' ORDER BY tBoxart_BoxFram DESC"
      End If
      
      RS_Open 1, SQL, False
          
        If rsDB(1).EOF Then
          hitID = 0
        Else
          hitID = rsDB(1)("tID")
        End If
       
      RS_Close 1
    Else
      hitID = 0
    End If
    
    getRegKod = bSure & "-" & hitID
  End Function
  
  Function ImgOriginal(lID)
    RS_Open 1, "SELECT * FROM cms_Bild WHERE bID = " & CLng(lID), False
    
      If rsDB(1).EOF Then
        ImgOriginal = "NO_IMG"
      Else
        ImgOriginal = "img_" & Right("0000000000" & lID, 10) & "_original." & rsDB(1)("bTyp")
      End If
    
    RS_Close 1
  End Function
  
  Function ImgDoRenew(myID, sSize)
    sSizes = Con.ExeCute("SELECT bInSizes FROM cms_Bild WHERE bID = " & CLng(myID))(0)
    
    If InStr(sSizes, sSize) > 0 Then
    Else
      mSize = Split(sSize, ",")
      ImgResize myID, mSize(0), mSize(1), 80
      
      sSizes = sSizes & ";" & sSize
      Con.ExeCute("UPDATE cms_Bild Set bInSizes = '" & sSizes & "' WHERE bID = " & CLng(myID))
    End If
  End Function
  
  Function ImgResize(sImgID, lWidth, lHeight, lCompression)
    sImage = config_ImageFolder & ImgOriginal(sImgID)
  
    Set Jpeg = Server.CreateObject("Persits.Jpeg")
      Jpeg.Open CStr(sImage)
      Jpeg.Canvas.Brush.Color = &HFFFFFF
      Jpeg.Interpolation = 10
      Jpeg.Quality = lCompression
      Jpeg.Progressive = True
      Jpeg.PNGOutput = True
      
      oWidth  = Jpeg.OriginalWidth
      oHeight = Jpeg.OriginalHeight
    
      Jpeg.PreserveAspectRatio = True
      
      nWidth_Diff = oWidth - lWidth
      nHeight_Diff = oHeight - lHeight
      
      If nWidth_Diff > nHeight_Diff Then
        If nWidth_Diff > -1 Then Jpeg.Width = lWidth
      Else
        If nHeight_Diff > -1 Then Jpeg.Height = lHeight
      End If
      
      nWidth  = Jpeg.Width
      nHeight = Jpeg.Height
      
      If nWidth < lWidth Then
        nx0 = -((lWidth - nWidth) / 2)
        nx1 = nWidth + ((lWidth - nWidth) / 2)
      Else
        nx0 = 0
        nx1 = nWidth
      End If
      
      If nHeight < lHeight Then
        ny0 = -((lHeight - nHeight) / 2)
        ny1 = nHeight + ((lHeight - nHeight) / 2)
      Else
        ny0 = 0
        ny1 = nHeight
      End If
      
      Jpeg.Crop nx0, ny0, nx1, ny1
      Jpeg.Crop 0, 0, lWidth, lHeight
      
      sFileSave = config_ImageFolder & lWidth & "x" & lHeight & "\img_" &  Right("0000000000" & sImgID, 10) & ".png"
      
      Jpeg.Save sFileSave
    Set Jpeg = Nothing
    
    ImgResize = sFileSave
  End Function
  
  Function getSpelTitel(lID)
    RS_Open 1, "SELECT * FROM cms_SpelTitlar WHERE tSpelID = " & CLng(lID) & " AND tID IN (SELECT sStandard_Titel FROM cms_Spel WHERE sID = " & CLng(lID) & ")" , False
    
      If rsDB(1).EOF Then
        retTitel = "-"
      Else
        retTitel = rsDB(1)("tTitel")
      End If
    
    RS_Close 1
    
    getSpelTitel = retTitel
  End Function
  
  Function getUserName(lID)
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID) , False
    
      If rsDB(1).EOF Then
        retTitel = "-"
      Else
        retTitel = rsDB(1)("aAnvNamn")
      End If
    
    RS_Close 1
    
    getUserName = retTitel
  End Function
  
  Function GetAcc(sDemand)
    bAccess = False
  
    If CONST_CMS_RIGHTS <> Empty Then
      If InStr(1, CONST_CMS_RIGHTS, sDemand, vbTextCompare) > 0 Then bAccess = True
    End If
    
    GetAcc = bAccess
  End Function
  
  Function PosterRights(sMethod, lForum, lID)
    haveAccess = True
  
    Select Case UCase(sMethod)
      Case "THREAD"
        If Not GetForumRights(lForum, "NewThread") Then haveAccess = False
      Case "ANSWER"
        If Not GetForumRights(lForum, "NewReply") Then haveAccess = False
      Case "THREAD_EDIT"
      Case "ANSWER_EDIT"
    End Select
    
    PosterRights = haveAccess
  End Function
  
  Function SayMe(sTitel,sText,sLank)
    Session.Value("trans_Titel") = sTitel
    Session.Value("trans_Text")  = sText
    Session.Value("trans_Lank")  = sLank
    
    Response.Redirect("/set/meddelande.asp")
  End Function
  
  Function TraceHyperlinks(sText)
    Dim RegExp
    Set RegExp = New RegExp

      RegExp.Global = True
      RegExp.IgnoreCase = True
      
      ' Hitta bilder (WWW)
      RegExp.Pattern = "([.,?!:;-]|\s)(www.)(.*?)(\.)(gif|png|jpg|jpeg)"
      sText = RegExp.Replace(sText & " ", "$1[img]$2$3$4$5[/img]")
      
      ' Hitta bilder (Utan vettig start)
      RegExp.Pattern = "(\s)(.*?)(\.)(gif|png|jpg|jpeg)(\s|[.,?!:;-])"
      sText = RegExp.Replace(sText & " ", "$1[img]$2$3$4[/img]$5")
      
      ' Hitta bilder (PNG, JPEG, JPG och GIF)
      RegExp.Pattern = "([.,?!:;-]|\s)(http|https|ftp)(://)(.*?)(\.)(gif|png|jpg|jpeg)"
      sText = RegExp.Replace(sText & " ", "$1[img]$2$3$4$5$6[/img]")
      
      ' Hitta länkar (WWW)
      RegExp.Pattern = "([.,?!:;-]|\s)(www.)(.*?)([.,?!:;-]\s|\s)"
      sText = RegExp.Replace(sText & " ", "$1[url=http://$2$3]$2$3[/url]$4")
      
      ' Hitta länkar (HTTP, HTTPS och FTP)
      RegExp.Pattern = "([.,?!:;-]|\s)(http|https|ftp)(://)(.*?)([.,?!:;-]\s|\s)"
      sText = RegExp.Replace(sText & " ", "$1[url]$2$3$4[/url]$5")

    Set RegExp = Nothing
    TraceHyperlinks = sText
  End Function
  
  Function MakeLegal(ByVal sText)
    For t = 1 To Len(sText)
      tkn = Mid(sText, t, 1)
      Select Case Asc(tkn)
        Case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90                        ' A-Z
          nText = nText & tkn
        Case 97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122 ' a-z
          nText = nText & tkn
        Case 48,49,50,51,52,53,54,55,56,57                                                                        ' 0-9
          nText = nText & tkn
        Case 229,228,246,197,196,214                                                                              ' åäöÅÄÖ
          nText = nText & tkn
        Case 95,45,35,94,32,38,44,46,58,64,128                                                                    ' _ - # ^ (space) & , . : @ €
          nText = nText & tkn
        Case 233,201,232,200,225,193,224,192,252,220,237,205,236,204,243,211,242,210                              ' éÉèÈáÁàÀüÜíÍìÌóÓòÒ
          nText = nText & tkn
      End Select
    Next
    
    MakeLegal = nText
  End Function
  
  Function MinutesToSplit(ByVal lMinutes)
    Dim minRest, giveNo
    giveNo = ""
    minRest = lMinutes
    
    lYears = RoundDown(minRest / 525600)
    minRest = minRest - (lYears * 525600)
    If lYears > 0 Then giveNo = giveNo & lYears & " år "
    
    lMonths = RoundDown(minRest / 40320)
    minRest = minRest - (lMonths * 40320)
    If lMonths > 0 Then giveNo = giveNo & lMonths & " månad(er) "
    
    lWeeks = RoundDown(minRest / 10080)
    minRest = minRest - (lWeeks * 10080)
    If lWeeks > 0 Then giveNo = giveNo & lWeeks & " veck(a/or) "
    
    lDays = RoundDown(minRest / 1440)
    minRest = minRest - (lDays * 1440)
    If lDays > 0 Then giveNo = giveNo & lDays & " dag(ar) | "
    
    lHours = RoundDown(minRest / 60)
    minRest = minRest - (lHours * 60)
    If lHours > 0 Then giveNo = giveNo & lHours & "h "
    
    If minRest > 0 Then giveNo = giveNo & minRest & "m "
    If giveNo = Empty Then giveNo = "0m"
    
    MinutesToSplit = giveNo
  End Function
  
  Function Text2Num(sText)
    If IsNumeric(sText) Then
      lNum = CLng(sText)
    Else
      sText = LCase(Trim(sText))
      Select Case sText 
        Case "twelve","tolv"                              : lNum = 12
        Case "eleven","elva"                              : lNum = 11
        Case "ten","tio","tie"                            : lNum = 10
        Case "nine","nio","nie"                           : lNum = 9
        Case "eight","åtta","åta","atta","aotta","otta"   : lNum = 8
        Case "seven","sju","su"                           : lNum = 7
        Case "sex","six","seks"                           : lNum = 6
        Case "five","fem"                                 : lNum = 5
        Case "four","fyra","fyr"                          : lNum = 4
        Case "three","tre"                                : lNum = 3
        Case "two","två","tva","tvao","tvo"               : lNum = 2
        Case "one","ett","en"                             : lNum = 1
        Case Else                                         : lNum = -1
      End Select
    End If
    
    Text2Num = CLng(lNum)
  End Function
  
  Function RemoveDoubleScore(ByVal sText)
    RemoveDoubleScore = Replace(sText, "--", "")
  End Function
  
  Function FixAlfaList(sAlfa)
    sAlfa = Trim(sAlfa & " ")
    If Len(sAlfa) > 0 Then
      Select Case UCase(sAlfa)
        Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "Å", "Ä", "Ö"
          nAlfa = UCase(sAlfa)
        Case "NUM"
          nAlfa = "NUM"
        Case Else 
          nAlfa = ""
      End Select
    Else
      nAlfa = ""
    End If
    
    FixAlfaList = nAlfa
  End Function
  
  Function AlfaToSQL(sAlfa, sField)
    Select Case sAlfa
      Case "NUM"
        nSQL = "AND " & sField & " LIKE '[0-9_^#.,:;-]_%'"
      Case ""
        nSQL = ""
      Case Else
        nSQL = "AND " & sField & " LIKE '" & sAlfa & "%'"
    End Select
    
    AlfaToSQL = nSQL
  End Function
  
  Function TitleToSQL(sTitle, sField)
    If Len(Trim(sTitle & " ")) > 0 Then
      nSQL = "AND " & sField & " LIKE '%" & sTitle & "%'"
    Else
      nSQL = ""
    End If
    
    TitleToSQL = nSQL
  End Function
  
  Function sEncode(sText)
    sEncode = Trim(Server.HTMLEncode(sText & " "))
  End Function
  
  Function SlumpText(lLength, bLimit)
  
    If bLimit Then
      sVals = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,1,2,3,4,5,6,7,8,9,0"
    Else
      sVals = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,1,2,3,4,5,6,7,8,9,0,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,-,#,?,!,@"
    End If
    
    sArrs = Split(sVals, ",")
    
    Randomize
    For zy = 1 To lLength
      lSlumpKey = CLng(Rnd*UBound(sArrs))
      nStr = nStr & sArrs(lSlumpKey)
    Next
    
    SlumpText = nStr
  End Function

  Function FixNum(lNum)
    If IsNull(lNum) Then
      newNum = 0
    Else
      newNum = CLng(lNum)
    End if
    
    FixNum = newNum
  End Function
  
  Function CutText(sText, lLength)
    sText = Trim(sText & " ")
    If Len(sText) > lLength Then sText = Left(sText, lLength-3) & "..."
    CutText = sText
  End Function
  
  Function FixSrc(sText)
    If Trim(sText & " ") <> Empty Then
      zx = 0
    
      sText = Trim(sText)
      sText = Replace(sText, ";;", "")
      sText = Replace(sText, ";|;", "")
      
      If Left(sText, 1) = ";" Then sText = Right(sText, Len(sText)-1)
      If Right(sText, 1) = ";" Then sText = Left(sText, Len(sText)-1)
    
      sKallor = Split(sText, ";")
      ReDim vKalla(UBound(sKallor))
      
      For Each sKalla In sKallor
        nDel = Split(sKalla, "|")
      
        If UBound(nDel) > 0 Then
          If LCase(Left(nDel(1), 7)) <> "http://" Then nDel(1) = "http://" & nDel(1)
          vKalla(zx) = "<a href='" & nDel(1) & "' rel='nofollow' target='_blank'>" & nDel(0) & "</a>"
          zx = zx + 1
        ElseIf UBound(nDel) > -1 Then
          vKalla(zx) = nDel(0)
          zx = zx + 1
        End If
      Next
      
      For zx = 0 To UBound(vKalla)
        If zx = UBound(vKalla) Then
          nText = nText & " och " & vKalla(zx)
        Else
          nText = nText & ", " & vKalla(zx)
        End If
      Next
      
      nText = Trim(nText)
      If Left(nText, 2) = ", " Then nText = Right(nText, Len(nText)-2)
      If Left(nText, 4) = "och " Then nText = Right(nText, Len(nText)-4)
      nText = Trim(nText)
    Else
      nText = "- Ingen källa -"
    End If
  
    FixSrc = nText
  End Function
  
  Function RoundUp(lNo1, lNo2)
    lSum = CDbl(CDbl(lNo1) / CDbl(lNo2))
    If Round(lSum) < lSum Then lSum = Round(lSum) + 1
    
    RoundUp = Round(lSum)
  End Function
  
  Function RoundDown(lValue)
    Dim nValue
    nValue = Round(lValue)
    If nValue > lValue Then nValue = nValue - 1
    
    RoundDown = nValue
  End Function
  
  Function IsInArray(sArray, sValue)
    retVal = False
  
    For Each sPost In sArray
      If CStr(sPost) = Cstr(sValue) Then retVal = True
    Next
    
    IsInArray = retVal
  End Function
  
  Function ArrayToQuery(sArray, sKey)
    For Each sPost In sArray
      nQuery = nQuery & "&amp;" & sKey & "=" & sPost
    Next
    
    ArrayToQuery = nQuery
  End Function
  
  Function ArrayToHidden(sArray, sKey)
    For Each sPost In sArray
      nHidden = nHidden & "<input type='hidden' name='" & sKey & "' value='" & sPost & "'>"
    Next
    
    ArrayToHidden = nHidden
  End Function
  
  Dim pagingOnPage, pagingNumOfPages, pagingNumOfPosts, pagingBOF, pagingEOF, pagingPages
  Function CreatePaging(lPostsPerPage,lNumOfPosts)
    pagingOnPage      = GetQ("page", "123", 0)
    pagingNumOfPosts  = lNumOfPosts + 1
    pagingNumOfPages  = RoundUp(pagingNumOfPosts, lPostsPerPage)
    
    If pagingOnPage < 1 Then pagingOnPage = 1
    If pagingOnPage > pagingNumOfPages Then pagingOnPage = pagingNumOfPages
    
    pagingBOF = ((pagingOnPage * lPostsPerPage) - lPostsPerPage)
    pagingEOF = (pagingBOF + lPostsPerPage) - 1
    If pagingEOF => pagingNumOfPosts Then pagingEOF = pagingNumOfPosts - 1
  End Function
  
  Function CreatePagingChooser()
    If pagingNumOfPages < 16 Then
      For zz = 1 To pagingNumOfPages
        nPages = nPages & zz & ","
      Next
    Else
      nFirstNum = pagingOnPage - 3
      nLastNum = pagingOnPage + 3
    
      If nFirstNum > 4 Then
        For zz = 1 To 3
          nPages = nPages & zz & ","
        Next
        nPages = nPages & "...,"
      Else
        If nFirstNum - 1 > 0 Then
          For zz = 1 To nFirstNum - 1
            nPages = nPages & zz & ","
          Next
        End If
      End If
      
      For zz = 3 To 1 Step -1
        If pagingOnPage - zz > 0 Then
          nPages = nPages & pagingOnPage - zz & ","
        End If
      Next
      
      nPages = nPages & pagingOnPage & ","
      
      For zz = 1 To 3
        If pagingOnPage + zz < pagingNumOfPages + 1 Then
          nPages = nPages & pagingOnPage + zz & ","
        End If
      Next
      
      If nLastNum < pagingNumOfPages - 3 Then
        nPages = nPages & "...,"
        For zz = pagingNumOfPages - 2 To pagingNumOfPages
          nPages = nPages & zz & ","
        Next
      Else
        If nLastNum + 1 < pagingNumOfPages + 1 Then
          For zz = nLastNum + 1 To pagingNumOfPages
            nPages = nPages & zz & ","
          Next
        End If
      End If 
    End If
    
    If Len(nPages) > 0 Then nPages = Left(nPages, Len(nPages)-1)
    pagingPages = Split(nPages, ",")
  End Function
  
  Function IsNothing(sText)
    sText = Trim(sText & " ")
    If Len(sText) > 0 Then
      IsNothing = False
    Else
      IsNothing = True
    End If
  End Function
  
  Function MonthName(dMonth)
    Select Case dMonth
      Case  1 : sRet = "Januari"
      Case  2 : sRet = "Februari"
      Case  3 : sRet = "Mars"
      Case  4 : sRet = "April"
      Case  5 : sRet = "Maj"
      Case  6 : sRet = "Juni"
      Case  7 : sRet = "Juli"
      Case  8 : sRet = "Augusti"
      Case  9 : sRet = "September"
      Case 10 : sRet = "Oktober"
      Case 11 : sRet = "November"
      Case 12 : sRet = "December"
    End Select
    
    MonthName = sRet
  End Function
  
  Function DatumReplace(dDatum)
    If IsDate(dDatum) Then
      lDiff = DateDiff("d", dDatum, Now, 2, 3)
      
      Select Case lDiff
        Case -1
          sText = "Imorgon " & FormatDateTime(dDatum, vbShortTime)
        Case 0
          sText = "Idag " & FormatDateTime(dDatum, vbShortTime)
        Case 1
          sText = "Igår " & FormatDateTime(dDatum, vbShortTime)
        Case 2
          sText = "Förrgår " & FormatDateTime(dDatum, vbShortTime)
        Case Else
          sText = FormatDateTime(dDatum, vbShortDate) & " " & FormatDateTime(dDatum, vbShortTime)
      End Select
    Else
      sText = "- Ogiltigt Datum -"
    End If
    
    DatumReplace = sText
  End Function
  
  Function RelDatum(dDatum)
    Select Case CStr(dDatum)
      Case "0", "N/A"   : rDatum = "N/A"
      Case "N/R"        : rDatum = "N/R"
      Case Else
        If IsDate(dDatum) Then
          rDatum = CStr(FormatDateTime(dDatum, vbShortDate))
        Else
          rDatum = CStr(dDatum)
        End If
    End Select
    
    RelDatum = rDatum
  End Function
  
  Function GetLastDayOfMonth(bYear,bMonth)
    If IsDate(bYear & "-" & Right("00" & bMonth, 2) & "-31") Then
      GetLastDayOfMonth = 31
    ElseIf IsDate(bYear & "-" & Right("00" & bMonth, 2) & "-30") Then
      GetLastDayOfMonth = 30
    ElseIf IsDate(bYear & "-" & Right("00" & bMonth, 2) & "-29") Then
      GetLastDayOfMonth = 29
    ElseIf IsDate(bYear & "-" & Right("00" & bMonth, 2) & "-28") Then
      GetLastDayOfMonth = 28
    End If
  End Function
  
  Function SelectAllRegions(selVal)
    RS_Open 9, "SELECT * FROM cms_Region ORDER BY rHighLight DESC, rNamn ASC", False
    
      Do until rsDB(9).EOF
        If CLng(selVal) = CLng(rsDB(9)("rID")) Then
          sOpts = sOpts & "<option value=""" & rsDB(9)("rID") & """ selected>" & rsDB(9)("rNamn") & "</option>"
        Else
          sOpts = sOpts & "<option value=""" & rsDB(9)("rID") & """>" & rsDB(9)("rNamn") & "</option>"
        End if
        rsDB(9).MoveNext
      Loop
    
    RS_Close 9
    
    Response.Write sOpts
  End Function
  
  Function GetRegion(lRegion)
    RS_Open 9, "SELECT * FROM cms_Region WHERE rID = " & CLng(lRegion), False
    
      If rsDB(9).EOF Then
        retRegion = "Region saknas"
      Else
        retRegion = rsDB(9)("rNamn")
      End If
    
    RS_Close 9
    
    GetRegion = retRegion
  End Function
  
  Function FormIDToArray(sForm)
    sForm = LCase(sForm)
  
    For Each nObj In Request.Form
      If LCase(Left(nObj,Len(sForm))) = sForm Then nVal = nVal & Right(nObj, Len(nObj)-Len(sForm)) & " ,"
    Next
    
    If Len(nVal) > 0 Then nVal = Left(nVal, Len(nVal)-1)
    
    FormIDToArray = Split(nVal,",")
  End Function
  
  Function ActivePage()
    qOnPage_Addr   = Request.ServerVariables("SERVER_NAME")
    qOnPage_Script = Request.ServerVariables("URL")
    qOnPage_Query  = Request.ServerVariables("QUERY_STRING")
    
    ActivePage = Server.HTMLEncode("http://" & qOnPage_Addr & qOnPage_Script & "?" &  qOnPage_Query)
  End Function
  
  Function HasAcc(sRatter,sDemand)
    bAccess = False
  
    If sRatter <> Empty Then
      If InStr(1, sRatter, sDemand, vbTextCompare) > 0 Then bAccess = True
    End If
    
    HasAcc = bAccess
  End Function
  
  ' ### GLOBAL DATA CACHE ###
    Function Cache_Create(sName, aData, bAny, TTL)
      Application.Lock
        Application("cache_" & sName & "_Data") = aData
        Application("cache_" & sName & "_Any") = bAny
        Application("cache_" & sName & "_TTL") = DateAdd("s", TTL, Now)
      Application.Unlock
    End Function
    
    Function Cache_Exist(sName)
      If Len(Application("cache_" & sName & "_TTL")) > 0 AND Application("cache_" & sName & "_TTL") > Now Then
        Cache_Exist = True
      Else
        Cache_Exist = False
      End if
    End Function
    
    Function Cache_Any(sName)
      Cache_Any = Application("cache_" & sName & "_Any")
    End Function
    
    Function Cache_Fetch(sName)
      Cache_Fetch = Application("cache_" & sName & "_Data")
    End Function
  ' ##################
  
  ' ### LOCAL USER CACHE ###
    Function user_Cache_Create(sName, aData, bAny, TTL)
      Session("user_cache_" & sName & "_Data") = aData
      Session("user_cache_" & sName & "_Any") = bAny
      Session("user_cache_" & sName & "_TTL") = DateAdd("s", TTL, Now)
    End Function
    
    Function user_Cache_Exist(sName)
      If Len(Session("user_cache_" & sName & "_TTL")) > 0 AND Session("user_cache_" & sName & "_TTL") > Now Then
        user_Cache_Exist = True
      Else
        user_Cache_Exist = False
      End if
    End Function
    
    Function user_Cache_Any(sName)
      user_Cache_Any = Session("user_cache_" & sName & "_Any")
    End Function
    
    Function user_Cache_Fetch(sName)
      user_Cache_Fetch = Session("user_cache_" & sName & "_Data")
    End Function
  ' ##################
  
  Function getColor(zc) ' ANNONSER
    Select Case zc
      Case 1    : getColor = "#962b20"
      Case 2    : getColor = "#20962e"
      Case 3    : getColor = "#28a49a"
      Case 4    : getColor = "#8c8b1d"
      Case 5    : getColor = "#818181"
      Case Else : getColor = "#3c88a8"
    End Select
  End Function
  
  Function SendMyMail(ByVal sText, ByVal sAmne, ByVal sEpost)
    Select Case UCase(KOMP_MAIL)
      Case "CDONTS"
        Set mMail = CreateObject("CDONTS.NewMail")
          mMail.From       = MAIL_NOREPLY
          mMail.To         = sEpost
          mMail.Subject    = sAmne
          mMail.BodyFormat = 0
          mMail.MailFormat = 0
          mMail.Body       = sText
          mMail.Send
        Set mMail = Nothing
      Case "CDOSYS"
        Set mMail = CreateObject("CDO.Message")
          sch = "http://schemas.microsoft.com/cdo/configuration/"
          Set cdoConfig = CreateObject("CDO.Configuration")

            cdoConfig.Fields.Item(sch & "sendusing") = 2
            cdoConfig.Fields.Item(sch & "smtpserver") = MAIL_SMTP
            cdoConfig.Fields.update
   
            Set mMail.Configuration = cdoConfig

            mMail.From       = MAIL_NOREPLY
            mMail.To         = sEpost
            mMail.Subject    = sAmne
            mMail.HTMLBody   = sText
            mMail.Send
            
          Set cdoConfig = Nothing 
        Set mMail = Nothing
      Case "JMAIL"
        Set mMail = Server.CreateObject("JMail.SMTPMail")
          mMail.ServerAddress = MAIL_SMTP
          mMail.Sender        = MAIL_NOREPLY
          mMail.Subject       = sAmne
          mMail.AddRecipient    sEpost
          mMail.HTMLBody      = sText
          mMail.Priority = 1
          mMail.Execute
        Set mMail = Nothing
      Case "ASPSMARTMAIL"
        Set mMail = Server.CreateObject("aspSmartMail.SmartMail")
          mMail.Server          = MAIL_SMTP
          mMail.SenderName      = MAIL_NAME
          mMail.SenderAddress   = MAIL_NOREPLY
          mMail.Recipients.Add    sEpost, sEpost
          mMail.Subject         = sAmne
          mMail.Body            = sText
          mMail.SendMail
        Set mMail = Nothing
      Case "ASPEMAIL"
        Set mMail = Server.CreateObject("Persits.MailSender") 
          mMail.Host        = MAIL_SMTP
          mMail.FromName    = MAIL_NAME
          mMail.From        = MAIL_NOREPLY
          mMail.AddAddress    sEpost, sEpost
          mMail.Subject     = sAmne
          mMail.Body        = sText
          mMail.Send
        Set mMail = Nothing
      Case "ASPMAIL"
        Set mMail = Server.CreateObject("SMTPsvg.Mailer")
          mMail.RemoteHost    = MAIL_SMTP
          mMail.FromName      = MAIL_NAME
          mMail.FromAddress   = MAIL_NOREPLY
          mMail.AddRecipient    sEpost, sEpost
          mMail.Subject       = sAmne
          mMail.BodyText      = sText
          mMail.SendMail
        Set mMail = Nothing
    End Select
  End FUnction
%>