<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>
<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

  <%
    Call start_Rec2Session("post")
    
      lID     = GetF("tradID","123",0)
      lIDSvar = GetF("tradID_Svar","123",0)
      sURL    = GetF("url", "ABC", 255)
      sURL    = Server.UrlEncode(sURL)

      If lID = 0 And lIDSvar = 0  Then lPostStat = 0    ' En helt ny tråd
      If lID = 0 And lIDSvar <> 0 Then lPostStat = 1    ' Ett nytt inlägg
      If lID <> 0 And lIDSvar = 0 Then lPostStat = 2    ' En tråd/inlägg redigeras
      
      sAmne       = GetF("fAmne","ABC",100)
      sTextM      = GetF("fMsg","ABC",20000)
      bLocked     = GetF("fLocked","CHK",0)
      bKlistrad   = GetF("fKlistrad","CHK",0)
      bDold       = GetF("fDold","CHK",0)
      bAutoUrl    = GetF("fAutoUrl","CHK",0)
      bAutoSmil   = GetF("fAutoSmil","CHK",0)
      dForum      = GetF("kategori","123",0)
      'bOwner      = GetF("fowner","CHK",0)
      
      If lPostStat = 2 Then
        lTradTyp    = IsMainThread(lID)
        If lTradTyp Then lPostStat = 2 Else lPostStat = 3
      End If
      
      If Len(Trim(sAmne)) < 1 Then Response.Redirect("../ny_trad.asp?fail=1&url=" & sURL)
      If Len(Trim(sTextM)) < 1 Then Response.Redirect("../ny_trad.asp?fail=2&url=" & sURL)
      
      Select Case lPostStat
        Case 0  ' En helt ny tråd
          lForum = dForum
          If Not dbForumExists(lForum) Then Response.Redirect("../ny_trad.asp?fail=4&url=" & sURL)
          If Not PosterRights("THREAD", lForum, 0) Then Response.Redirect("../ny_trad.asp?fail=3&A=1&url=" & sURL)
          
          bNew         = True
          bStatus_Trad = True
          lSvarsId     = 0
          bMovable     = True
          bNewTitle    = True
          bEdited      = False
          bMainT       = True
          
          'GetRightsForum forumID ' Hämta fram rättigheterna
          'If Not sec_Trad_Skapa Then Response.Redirect("../ny_trad.asp?a=" & lIDSvar & "&fail=3&A=4&url=" & sURL)
          
          
          nLogg = nLogg & Now & " | [Tråd] Skapad - Av [" & CONST_USERID & "] " & CONST_USERNAME
        Case 1  ' Ett nytt inlägg
          lForum = GetForumFromID(lIDSvar)
          If lForum = 0                            Then Response.Redirect("../ny_trad.asp?fail=3&A=2&url=" & sURL)
          If Not PosterRights("ANSWER", lForum, 0) Then Response.Redirect("../ny_trad.asp?fail=3&A=3&url=" & sURL)
          
          bNew         = True
          bStatus_Trad = False
          lSvarsId     = lIDSvar
          bMovable     = True
          bNewTitle    = True
          bEdited      = False
          bMainT       = False
          
          GetRights lIDSvar ' Hämta fram rättigheterna
          If Not sec_Inlagg_Skapa Then Response.Redirect("../ny_trad.asp?a=" & lIDSvar & "&fail=3&A=4&url=" & sURL)
          
          nLogg = nLogg & Now & " | [Svar] Skapad - Av [" & CONST_USERID & "] " & CONST_USERNAME
        Case 2  ' En tråd redigeras
          lForum = dForum
          If Not dbForumExists(lForum) Then Response.Redirect("../ny_trad.asp?fail=4&url=" & sURL)
          If Not PosterRights("THREAD_EDIT", lForum, lID) Then Response.Redirect("../ny_trad.asp?fail=3&A=5&url=" & sURL)
          
          bNew         = False
          bMovable     = False
          bNewTitle    = False
          bEdited      = True
          bMainT       = True
          
          GetRights lID ' Hämta fram rättigheterna
          If Not sec_Trad_Hantera Then Response.Redirect("../ny_trad.asp?e=" & lID & "&fail=3&A=6&url=" & sURL)
          
          nLogg = nLogg & ";" & Now & " | [Tråd] Ändrad - Av [" & CONST_USERID & "] " & CONST_USERNAME
        Case 3  ' Ett inlägg redigeras
          lForum = GetForumFromID(lID)
          If Not PosterRights("ANSWER_EDIT", lForum, lID) Then Response.Redirect("../ny_trad.asp?fail=3&A=7&url=" & sURL)
          
          bNew         = False
          bMovable     = True
          bNewTitle    = True
          bEdited      = True
          bMainT       = False
          
          GetRights lID ' Hämta fram rättigheterna
          GetUserStatsFromPost lID
          If Not sec_Hantera(lPost_UserAdmin, lPost_UserID) Then Response.Redirect("../ny_trad.asp?e=" & lID & "&fail=3&A=8&url=" & sURL)
          
          nLogg = nLogg & ";" & Now & " | [Svar] Ändrad - Av [" & CONST_USERID & "] " & CONST_USERNAME
        Case Else
          Response.Redirect("../ny_trad.asp?fail=uhm&url=" & sURL)
      End Select
    
      RS_Open 1, "SELECT * FROM fsBB_Tradar WHERE tID = " & CLng(lID), True
        
        If rsDB(1).EOF And bEdited Then Response.Redirect("../ny_trad.asp?fail=3&A=9&url=" & sURL)
        
        If bNew Then
          rsDB(1).AddNew
          
          bIsNew = True
          
          rsDB(1)("tDatum_Skapad")      = Now
          rsDB(1)("tDatum_Uppdaterad")  = Now
          rsDB(1)("tAnv_Skapad")        = CONST_USERID
          rsDB(1)("tAnv_Uppdaterad")    = CONST_USERID
          
          rsDB(1)("tStatus_Undertrad")  = lSvarsId
          rsDB(1)("tStatus_Trad")       = bStatus_Trad
          
          rsDB(1)("tSec_IP")            = Left(Request.ServerVariables("REMOTE_ADDR"), 40) 
          
          If lSvarsId > 0 Then Con.ExeCute("UPDATE fsBB_Tradar SET tANv_Uppdaterad = " & CLng(CONST_USERID) & ", tDatum_Uppdaterad = '" & Now & "' WHERE tID = " & CLng(lSvarsId))
        End If
        
        If bMovable Then
          rsDB(1)("tForum")         = lForum
        Else
          If sec_Trad_Admin And CLng(lForum) <> CLng(rsDB(1)("tForum")) And bMainT Then
            rsDB(1)("tForum")         = lForum
            Con.ExeCute("UPDATE fsBB_Tradar SET tForum = " & CLng(lForum) & " WHERE tStatus_Trad = 0 And tStatus_UnderTrad = " & CLng(lID))
            nLogg = nLogg & ";" & Now & " | [Tråd] Flyttad - Av [" & CONST_USERID & "] " & CONST_USERNAME
          End If
        End If
        
        If bNewTitle Then
          rsDB(1)("tAmne")          = sAmne
        Else
          If sec_Trad_Admin Then
            rsDB(1)("tAmne")          = sAmne
          End If
        End If
        
        If sec_Trad_Admin And bMainT Then
          If rsDB(1)("tStatus_Last") <> bLocked OR bIsNew Then
            rsDB(1)("tStatus_Last")     = bLocked
            If bLocked Then
              nLogg = nLogg & ";" & Now & " | [Tråd] Låst - Av [" & CONST_USERID & "] " & CONST_USERNAME
            Else
              nLogg = nLogg & ";" & Now & " | [Tråd] Upplåst - Av [" & CONST_USERID & "] " & CONST_USERNAME
            End If
          End If
          
          If rsDB(1)("tInst_Klistrad") <> bKlistrad OR bIsNew Then
            rsDB(1)("tInst_Klistrad")   = bKlistrad
            If bKlistrad Then
              nLogg = nLogg & ";" & Now & " | [Tråd] Klistrad - Av [" & CONST_USERID & "] " & CONST_USERNAME
            Else
              nLogg = nLogg & ";" & Now & " | [Tråd] Avklistrad - Av [" & CONST_USERID & "] " & CONST_USERNAME
            End If
          End If
          
          If rsDB(1)("tStatus_Dold") <> bDold OR bIsNew Then
            rsDB(1)("tStatus_Dold")     = bDold
            If bDold Then
              nLogg = nLogg & ";" & Now & " | [Tråd] Dold - Av [" & CONST_USERID & "] " & CONST_USERNAME
            Else
              nLogg = nLogg & ";" & Now & " | [Tråd] Synlig - Av [" & CONST_USERID & "] " & CONST_USERNAME
            End If
          End If
        End If
        
        rsDB(1)("tInst_Smilies")    = bAutoSmil
        rsDB(1)("tInst_Autolankar") = bAutoUrl
        
        If bAutoUrl Then sTextM = TraceHyperlinks(sTextM)
        rsDB(1)("tTextM")           = sTextM
        
        rsDB(1)("tDatum_Andrad")    = Now
        rsDB(1)("tAnv_Andrad")      = CONST_USERID
        
        rsDB(1)("tSec_ChIP")        = Left(Request.ServerVariables("REMOTE_ADDR"), 40) 
        
        rsDB(1)("tLogg")            = rsDB(1)("tLogg") & nLogg
        
        rsDB(1).Update
        
        Select Case lPostStat
          Case 0 : lReturnID = rsDB(1)("tID")               : lMark = rsDB(1)("tID")
          Case 1 : lReturnID = lIDSvar                      : lMark = rsDB(1)("tID")
          Case 2 : lReturnID = lID                          : lMark = rsDB(1)("tID")
          Case 3 : lReturnID = rsDB(1)("tStatus_UnderTrad") : lMark = rsDB(1)("tID")
        End Select
      
      RS_Close 1
    
    Call stop_Rec2Session("post")
    Call SayMe("Postad","Ditt <strong>foruminlägg</strong> har nu postats!", "/avdelning/forum/trad.asp?e=" & lReturnID & "&go2=" & lMark)

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->