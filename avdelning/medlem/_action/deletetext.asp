<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lID       = GetQ("e","123",0)
  sTextAvd  = UCase(GetQ("avd","ABC",10))
  
  filter_list   = GetQ("list","ABC",15)
  filter_page   = CLng(GetQ("page","123",0))
  
  Select Case sTextAvd
    Case "REC"
      RS_Open 1, "SELECT * FROM cms_Recensioner WHERE (rStatus = 1 Or rStatus = 3) And rSkapadAv = " & CLng(CONST_USERID) & " And rID = " & CLng(lID), True
        If Not rsDB(1).EOF Then
          rsDB(1)("rStatus") = 0
          rsDB(1).Update
        End If 
      RS_Close 1
    Case "ART"
      RS_Open 1, "SELECT * FROM cms_Artiklar WHERE (aaStatus = 1 Or aaStatus = 3) And aaSkapadAv = " & CLng(CONST_USERID) & " And aaID = " & CLng(lID), True
        If Not rsDB(1).EOF Then
          rsDB(1)("aaStatus") = 0
          rsDB(1).Update
        End If 
      RS_Close 1
    Case "TOT"
      RS_Open 1, "SELECT * FROM cms_SpelTrix WHERE (xStatus = 1 Or xStatus = 3) And xSkapadAv = " & CLng(CONST_USERID) & " And xID = " & CLng(lID), True
        If Not rsDB(1).EOF Then
          rsDB(1)("xStatus") = 0
          rsDB(1).Update
        End If 
      RS_Close 1
  End Select
  
  Response.Redirect("../minatexter.asp?list=" & filter_list & "&page=" & filter_page)

%>