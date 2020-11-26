<!--#INCLUDE FILE="../__INC/includes.asp"-->

  <%
    If CONST_LOGIN Then

      lAnvID    = CLng(CONST_USERID)
      lSpelID   = GetQ("e","123",0)
      lBetyg    = GetQ("b","123",0)
      If lBetyg < 1 Or lBetyg > 6 Then lBetyg = 0
      
      bSetBetyg = False
      
      If lBetyg > 0 And lSpelID > 0 Then
        RS_Open 1, "SELECT * FROM cms_Spel WHERE sID = " & CLng(lSpelID), False
          If Not rsDB(1).EOF Then bSetBetyg = True
        RS_Close 1
        
        If bSetBetyg Then
          RS_Open 1, "SELECT * FROM cms_SpelBetyg WHERE bAnv = " & CLng(lAnvID) & " AND bSpelID = " & CLng(lSpelID), True
            If rsDB(1).EOF Then
              rsDB(1).AddNew
              rsDB(1)("bAnv")          = lAnvID
              rsDB(1)("bSpelID")       = lSpelID
            End if
              
            rsDB(1)("bBetyg")          = lBetyg
            
            rsDB(1).Update
          RS_Close 1
        End If
      End If
    
    End If
  %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->