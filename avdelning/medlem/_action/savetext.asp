<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%
    
    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    Call start_Rec2Session("txt")
    
      lID       = GetF("e","123",0)

      sTextAvd      = UCase(GetF("avd","ABC",10))
  
      filter_list   = GetF("list","ABC",15)
      filter_page   = CLng(GetF("page","123",0))
      
      filter_all    = "&list=" & filter_list & "&page=" & filter_page & "&avd=" & sTextAvd
      
      sTextM    = GetF("TextM","ABC",20000)
      
      If Len(Trim(sTextM)) < 50 Then Response.Redirect("../redigeratext.asp?e=" & CLng(lID) & filter_all & "&fail=1")
    
      Select Case sTextAvd
        Case "REC"
          RS_Open 1, "SELECT * FROM cms_Recensioner WHERE (rStatus = 1 Or rStatus = 3) And rSkapadAv = " & CLng(CONST_USERID) & " And rID = " & CLng(lID), True
            If Not rsDB(1).EOF Then
              rsDB(1)("rText")    = sTextM
              If CONST_PUBLISH Then
                rsDB(1)("rStatus")          = 4
                rsDB(1)("rDatumPublicerad") = Now
                rsDB(1)("rPubliceradAv")    = CONST_USERID
              Else
                rsDB(1)("rStatus")  = 2
              End If
              rsDB(1).Update
            Else
              Response.Redirect("minatexter.asp?list=" & filter_list & "&page=" & filter_page)
            End If
          RS_Close 1
        Case "ART"
          RS_Open 1, "SELECT * FROM cms_Artiklar WHERE (aaStatus = 1 Or aaStatus = 3) And aaSkapadAv = " & CLng(CONST_USERID) & " And aaID = " & CLng(lID), True
            If Not rsDB(1).EOF Then
              rsDB(1)("aaText")    = sTextM
              If CONST_PUBLISH Then
                rsDB(1)("aaStatus")          = 4
                rsDB(1)("aaDatumPublicerad") = Now
                rsDB(1)("aaPubliceradAv")    = CONST_USERID
              Else
                rsDB(1)("aaStatus")  = 2
              End If
              rsDB(1).Update
            Else
              Response.Redirect("minatexter.asp?list=" & filter_list & "&page=" & filter_page)
            End If
          RS_Close 1
        Case "TOT"
          RS_Open 1, "SELECT * FROM cms_SpelTrix WHERE (xStatus = 1 Or xStatus = 3) And xSkapadAv = " & CLng(CONST_USERID) & " And xID = " & CLng(lID), True
            If Not rsDB(1).EOF Then
              rsDB(1)("xTextM")    = sTextM
              If CONST_PUBLISH Then
                rsDB(1)("xStatus")          = 4
                rsDB(1)("xDatumPublicerad") = Now
                rsDB(1)("xPubliceradAv")    = CONST_USERID
              Else
                rsDB(1)("xStatus")  = 2
              End If
              rsDB(1).Update
            Else
              Response.Redirect("minatexter.asp?list=" & filter_list & "&page=" & filter_page)
            End If
          RS_Close 1
        Case Else
          Response.Redirect("minatexter.asp?list=" & filter_list & "&page=" & filter_page)
      End Select
    
    Call stop_Rec2Session("txt")
    Call SayMe("Sparad","Din <strong>text</strong> har nu sparats, den kommer nu kontrolleras av oss innan den publiceras!", "/avdelning/medlem/minatexter.asp?list=" & filter_list & "&page=" & filter_page)

  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->