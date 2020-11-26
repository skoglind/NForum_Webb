<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lID       = GetQ("e","123",0)
  lMedlem   = GetQ("m","ABC",50)
  
  If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME

  If Not dbUserExists(lMedlem) Then Response.Redirect("/")
  anvID = GetIDFromUsername(lMedlem)
  
  If config_LockDown_Feedback Then Response.Redirect("../default.asp?m=" & lMedlem)
  
    If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
      RS_Open 1, "SELECT * FROM cms_Feedback WHERE fbID = " & CLng(lID), True
    Else
      RS_Open 1, "SELECT * FROM cms_Feedback WHERE fbAnv = " & CLng(CONST_USERID) & " And fbID = " & CLng(lID), True
    End If
      If Not rsDB(1).EOF Then
        rsDB(1)("fbRaderadAv") = CLng(CONST_USERID)
        rsDB(1).Update
      End If 
    RS_Close 1
  
  Response.Redirect("../omdome.asp?m=" & lMedlem)

%>