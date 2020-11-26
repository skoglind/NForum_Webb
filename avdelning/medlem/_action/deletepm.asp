<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lID   = GetQ("e","123",0)

  RS_Open 1, "SELECT * FROM fsBB_PM WHERE (pFran = " & CLng(CONST_USERID) & " OR pTill = " & CLng(CONST_USERID) & ") AND pID = " & CLng(lID), True
    
    If Not rsDB(1).EOF Then
      If CLng(rsDB(1)("pFran")) = CLng(CONST_USERID) Then
        lAvd = 1
        rsDB(1)("pRaderadFran") = True
      Else
        lAvd = 0
        rsDB(1)("pRaderadTill") = True
      End If
      
      rsDB(1).Update
    End If
    
  RS_Close 1
  
  If lAvd = 1 Then
    Response.Redirect("../skickat.asp")
  Else
    Response.Redirect("../inkorg.asp")
  End If

%>