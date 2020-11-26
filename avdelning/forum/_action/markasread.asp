<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If config_LockDown_Forum Then Response.Redirect("../default.asp")
  If Not CONST_LOGIN Then Response.Redirect("../default.asp")
  
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0

  nDate = DateAdd("d", -config_RemoOlasta, Now)
  
  lForum = GetQ("f","123",0)
  If lForum > 0 Then
    sqlForum = "AND fID = " & lForum & " "
  End If
  
  RS_Open 1, "SELECT tID, tDatum_Uppdaterad " & _
             "FROM fsBB_Tradar AS tbTrad " & _
             "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
             "WHERE tDatum_Uppdaterad > '" & nDate & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tStatus_Trad = 1 AND tStatus_Raderad = 0 " & _
             sqlForum & _
             "AND tID NOT IN (SELECT oTradID FROM fsBB_Olast WHERE oAnvandare = " & CLng(CONST_USERID) & ") ", False
    
      If rsDB(1).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
    If any_Tradar Then
      For zx = 0 To UBound(list_Tradar, 2)
        Con.ExeCute("INSERT INTO fsBB_Olast (oTradID, oDatum, oAnvandare) VALUES(" & CLng(list_Tradar(0,zx)) & ",'" & DateAdd("n", 1,Now) & "'," & CLng(CONST_USERID) & ")")
      Next
    End If

    
    If lForum > 0 Then
      Response.Redirect("../forum.asp?e=" & lForum)
    Else
      Response.Redirect("../default.asp")
    End If
%>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->