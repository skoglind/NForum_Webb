<%

  ' #### ÖVERSÄTTNINGSFUNKTIONER ####

    Function UserIDFromName(sUsrName)
      If Trim(sUsrName) = Empty Then sUsrName = ""
      sUsrName = MakeLegal(sUsrName)
    
      RS_Open 9, "SELECT * FROM fsBB_Anv WHERE aAnvNamn = '" & sUsrName & "'", False
      
        If Not rsDB(9).EOF Then
          retID = CLng(rsDB(9)("aID"))
        Else
          retID = 0
        End If
      
      RS_Close 9
      
      UserIDFromName = retID
    End Function
    
    Function GetIDFromUsername(sAnvNamn)
      sAnvNamn = MakeLegal(sAnvNamn)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE aAnvNamn = '" & sAnvNamn & "'")(0)
    
      If lCnt > 0 Then
        lAnvID = Con.ExeCute("SELECT aID FROM fsBB_Anv WHERE aAnvNamn = '" & sAnvNamn & "'")(0)
      Else
        lAnvID = 0
      End If
      
      GetIDFromUsername = lAnvID
    End Function
    
    Function GetForumFromID(lID)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tID = " & CLng(lID))(0)
    
      If lCnt > 0 Then
        lForumID = Con.ExeCute("SELECT tForum FROM fsBB_Tradar WHERE tID = " & CLng(lID))(0)
      Else
        lForumID = 0
      End If
      
      GetForumFromID = lForumID
    End Function
    
    Dim lPost_UserID, lPost_UserAdmin, lPost_TradID
    Function GetUserStatsFromPost(lID)
      'newID         = Con.ExeCute("SELECT tStatus_UnderTrad FROM fsBB_Tradar WHERE tID = " & CLng(lID))(0)
      'lPost_UserID  = Con.ExeCute("SELECT tAnv_Skapad FROM fsBB_Tradar WHERE tID = " & CLng(lID))(0)

      RS_Open 9, "SELECT * FROM fsBB_Tradar " & _
                 "LEFT JOIN fsBB_Anv ON tAnv_Skapad = fsBB_Anv.aID " & _
                 "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
                 "WHERE tID = " & CLng(lID), False
          
        If Not rsDB(9).EOF Then
          lPost_UserID    = rsDB(9)("tAnv_Skapad")
          lPost_TradID    = rsDB(9)("tID")
          lPost_UserAdmin = rsDB(9)("ttAdmin")
        Else
          lPost_UserID    = 0
          lPost_TradID    = 0
          lPost_UserAdmin = False
        End If
        
      RS_Close 9
    End Function
    
    Function GetSendPM(lID)
      bCanSend = True
    
      If lID = config_SystemUser Then bCanSend = False
      If bCanSend Then
        lCnt = Con.ExeCute("SELECT COUNT(aID) FROM fsBB_Anv WHERE aID = " & CLng(lID))(0)
        If lCnt > 0 Then
          If Not Con.ExeCute("SELECT aAktiveraPM FROM fsBB_Anv WHERE aID = " & CLng(lID))(0) Then bCanSend = False
        End If
        
        If CONST_ADMIN Then bCanSend = True 
      End If  
      
      GetSendPM = bCanSend
    End Function
  
  ' ################################
  
  ' #### EXISTERAR DET ####
  
    Function dbUserExists(sUsrName)
      sUsrName = Trim(MakeLegal(sUsrName))
      
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE aAnvNamn = '" & sUsrName & "'")(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbUserExists = bRet
    End Function
    
    Function dbForumExists(lForum)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Forum WHERE fID = " & CLng(lForum))(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbForumExists = bRet
    End Function
    
    Function dbTradExists(lTrad)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tStatus_Trad = 1 AND tID = " & CLng(lTrad))(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbTradExists = bRet
    End Function
    
    Function dbSpelExists(lSpel)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM cms_Spel WHERE sID = " & CLng(lSpel))(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbSpelExists = bRet
    End Function
    
    Function dbKonsolExists(lKonsol)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM cms_Konsol WHERE kID = " & CLng(lKonsol))(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbKonsolExists = bRet
    End Function
    
    Function dbTillbehorExists(lTillbehor)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM cms_Tillbehor WHERE iID = " & CLng(lTillbehor))(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbTillbehorExists = bRet
    End Function
    
    Function dbMailExists(sMail)
      sMail = Trim(MakeLegal(sMail))
      
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE aEpost = '" & sMail & "' OR aNyEpost = '" & sMail & "'")(0)
      If lCnt > 0 Then bRet = True Else bRet = False
      
      dbMailExists = bRet
    End Function
  
  ' #######################
  
  ' #### HÄMTA DATABASVÄRDEN ####
  
    Function GetForumRights(lKat, sMethod)
      haveAccess = True
      
      lNo = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Forum WHERE (fSec_" & sMethod & " = '0' OR fSec_" & sMethod & " LIKE '%;" & SEC_TITEL & ";%') AND fID = " & CLng(lKat))(0)
      If lNo < 1 Then haveAccess = False
      
      GetForumRights = haveAccess
    End Function
    
    Function GetNewPM()
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_PM WHERE pTill = " & CLng(CONST_USERID) & " AND pLast = 0")(0)

      GetNewPM = CLng(lCnt)
    End Function
    
    Function GetOnline()
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE aTimeStamp > '" & DateAdd("n", -5, Now) & "' AND aBlockadTill < '" & Date & "' AND aAktiverad = 1")(0)

      GetOnline = CLng(lCnt)
    End Function
  
  ' #############################
  
  ' #### ÖVRIGT ####
  
    Dim sec_Trad_Visa, sec_Trad_Skapa, sec_Trad_Hantera, sec_Trad_Admin
    Dim sec_Inlagg_Skapa, sec_Inlagg_Hantera, sec_Inlagg_Admin
    
    Function GetRights(valTrad)
      lTrad = valTrad
      noRights = False
    
      sec_Trad_Visa       = False
      sec_Trad_Skapa      = False
      sec_Trad_Hantera    = False
      sec_Trad_Admin      = False
      sec_Inlagg_Skapa    = False
      sec_Inlagg_Hantera  = 0
      sec_Inlagg_Admin    = 0
      
      bSkapaTrad          = False
      bSkapaInlagg        = False
      bVisa               = False
      bModeratorer        = False
    
      If CONST_LOGIN Then
        lUserID   = CONST_USERID
        lUserTT   = CONST_TITEL
        lUserAdm  = CONST_ADMIN
        
        RS_Open 9, "SELECT * FROM fsBB_Tradar WHERE tID = " & CLng(lTrad), False
          
          If Not rsDB(9).EOF Then
            If Not rsDB(9)("tStatus_Trad") Then lTrad = CLng(rsDB(9)("tStatus_UnderTrad"))
          Else
            noRights = True
          End If
        
        RS_Close 9
        
        If Not noRights Then
          RS_Open 9, "SELECT *, f.fSec_Mod AS tModerator, f.fSec_NewThread AS tSkapaTrad, f.fSec_NewReply AS tSkapaInlagg, f.fSec_View AS tVisa, t.ttAdmin AS aAdmin FROM fsBB_Tradar " & _
                     "LEFT JOIN fsBB_Forum AS f ON tForum = f.fID " & _
                     "LEFT JOIN fsBB_Anv AS a ON tAnv_Skapad = a.aID " & _
                     "LEFT JOIN fsBB_Titlar AS t ON a.aTitelID = t.ttID " & _
                     "WHERE tID = " & CLng(lTrad), False
                     
            If Not rsDB(9).EOF Then
              sSkapaTrad    = rsDB(9)("tSkapaTrad")
              sSkapaInlagg  = rsDB(9)("tSkapaInlagg")
              sVisa         = rsDB(9)("tVisa")
              sModeratorer  = rsDB(9)("tModerator")
              
              bLast         = rsDB(9)("tStatus_Last")
              bAdmin        = rsDB(9)("aAdmin")
              lUserTradID   = rsDB(9)("tAnv_Skapad")
              
              If InStr(sSkapaTrad, ";" & lUserTT & ";")   Or sSkapaTrad = "0"   Then bSkapaTrad = True
              If InStr(sSkapaInlagg, ";" & lUserTT & ";") Or sSkapaInlagg = "0" Then bSkapaInlagg = True
              If InStr(sVisa, ";" & lUserTT & ";")        Or sVisa = "0"        Then bVisa = True
              If InStr(sModeratorer, ";" & lUserTT & ";")                       Then bModeratorer = True
              
              ' ## FÅR HAN ENS SE TRÅDEN
              If bVisa Then
                sec_Trad_Visa = True
              End If
              
              ' ## FÅR HAN SKAPA TRÅDAR I FORUMET
              If bSkapaTrad Then
                sec_Trad_Skapa = True
              End If
              
              ' ## FÅR HAN SKAPA INLÄGG I TRÅDEN
              If bSkapaInlagg And Not bLast Then
                sec_Inlagg_Skapa = True
              ElseIf bSkapaInlagg And bLast And lUserAdm Then
                sec_Inlagg_Skapa = True
              ElseIf bSkapaInlagg And bLast And bModeratorer Then
                sec_Inlagg_Skapa = True
              End If
              
              ' ## FÅR HAN HANTERA TRÅDEN
              If lUserTradID = lUserID And Not bLast And bVisa Then 
                sec_Trad_Hantera = True
              ElseIf bModeratorer And bVisa And Not bAdmin Then
                sec_Trad_Hantera = True
              ElseIf lUserAdm And bVisa Then
                sec_Trad_Hantera = True
              End If
              
              ' ## FÅR HAN ADMINISTRERA TRÅDEN
              If bModeratorer And bVisa And Not bAdmin Then
                sec_Trad_Admin = True
              ElseIf lUserAdm And bVisa Then
                sec_Trad_Admin = True
              End If
              
              ' ## FÅR HAN HANTERA INLÄGG
              If lUserAdm And bVisa Then
                sec_Inlagg_Hantera = 3
              ElseIf bModeratorer And bVisa Then
                sec_Inlagg_Hantera = 2
              ElseIf Not bLast Then
                sec_Inlagg_Hantera = 1
              End If
              
              ' ## FÅR HAN ADMINISTRERA INLÄGG
              If lUserAdm And bVisa Then
                sec_Inlagg_Admin = 2
              ElseIf bModeratorer And bVisa Then
                sec_Inlagg_Admin = 1
              End If
              
            End If
            
          RS_Close 9
        End If
      Else
        RS_Open 9, "SELECT * FROM fsBB_Tradar LEFT JOIN fsBB_Forum ON fID = tForum WHERE tID = " & CLng(valTrad), False
          If Not rsDB(9).EOF Then
            If rsDB(9)("fSec_View") = "0" Then sec_Trad_Visa = True Else sec_Trad_Visa = False
          Else
            sec_Trad_Visa = False
          End If
        RS_Close 9
      End If
      
      ' ##
      '  sec_Trad_Skapa      > Skapa trådar                              [0-1] 
      '  sec_Trad_Hantera    > Redigera tråden                           [0-1] 
      '  sec_Trad_Admin      > Administrera tråden (låsa,flytta,radera)  [0-1]
      '  sec_Inlagg_Skapa    > Skapa inlägg                              [0-1]
      '  sec_Inlagg_Hantera  > Redigera inlägget                         [0-1-2-3] Där två även kan fibbla med andras inlägg
      '  sec_Inlagg_Admin    > Administrera inlägg (radera)              [0-1-2] Där två även kan röra andra admins inlägg
      ' ##
    End Function
    
    Function sec_Hantera(bAdmin, lID)
      lUserID   = CONST_USERID
      
      If sec_Inlagg_Hantera = 3 Then
        sec_Hantera = True
      ElseIf sec_Inlagg_Hantera = 2 And Not bAdmin Then
        sec_Hantera = True
      ElseIf sec_Inlagg_Hantera = 1 And CLng(lID) = CLng(lUserID) And Not bAdmin Then
        sec_Hantera = True
      Else
        sec_Hantera = False
      End If
    End Function
    
    Function sec_Admin(bAdmin)
      If sec_Inlagg_Admin = 2 Then
        sec_Admin = True
      ElseIf sec_Inlagg_Admin = 1 And Not bAdmin Then
        sec_Admin = True
      Else
        sec_Admin = False
      End If
    End Function
  
    Function IsMainThread(lID)
      lCnt = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Tradar WHERE tID = " & CLng(lID))(0)
    
      If lCnt > 0 Then
        bStatus = Con.ExeCute("SELECT tStatus_Trad FROM fsBB_Tradar WHERE tID = " & CLng(lID))(0)
        If bStatus Then bRet = True Else bRet = False
      Else
        bRet = False
      End If
      
      IsMainThread = bRet
    End Function
  
  ' ################

%>