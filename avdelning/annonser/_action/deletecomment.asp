<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If config_LockDown_Kommentarer Then Response.Redirect("../default.asp")
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lID       = GetQ("e","123",0)
  
    If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
      RS_Open 1, "SELECT * FROM cms_Kommentar_KopSalj WHERE kskID = " & CLng(lID), True
    Else
      RS_Open 1, "SELECT * FROM cms_Kommentar_KopSalj WHERE kskAnv = " & CLng(CONST_USERID) & " And kskID = " & CLng(lID), True
    End If
      If Not rsDB(1).EOF Then
        rsDB(1)("kskRaderadAv") = CLng(CONST_USERID)
        rsDB(1).Update
        lReturnID = rsDB(1)("kskAnnons")
      End If 
    RS_Close 1
  
  Response.Redirect("../annons_visa.asp?e=" & CLng(lReturnID))

%>