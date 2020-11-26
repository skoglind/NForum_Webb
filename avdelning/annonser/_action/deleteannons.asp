<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lID       = GetQ("e","123",0)
  
    doRadering = False
    If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
      RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksID = " & CLng(lID), True
    Else
      RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksSkapadAv = " & CLng(CONST_USERID) & " And ksID = " & CLng(lID), True
    End if
      If Not rsDB(1).EOF Then
        rsDB(1).Delete
        
        doRadering = True
      End If 
    RS_Close 1
  
    If doRadering Then
      Con.ExeCute("DELETE FROM cms_Kommentar_KopSalj WHERE kskAnnons = " & CLng(lID))
    End If
    
  Response.Redirect("../minaannonser.asp")

%>