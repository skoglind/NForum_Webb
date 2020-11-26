<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>
<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%
  lID = GetQ("e","123",0)
  retLink = "../default.asp"  

  RS_Open 1, "SELECT * FROM fsBB_Tradar " & _ 
             "LEFT JOIN fsBB_Anv ON fsBB_Anv.aID = tAnv_Skapad " & _
             "LEFT JOIN fsBB_Titlar ON fsBB_Titlar.ttID = aTitelID " & _
             "WHERE tID = " & CLng(lID), True
    
    If Not rsDB(1).EOF Then
      mainThread  = rsDB(1)("tStatus_Trad")
      userIsAdm   = rsDB(1)("ttAdmin")
      inForum     = rsDB(1)("tForum")
      inUTrad     = rsDB(1)("tStatus_UnderTrad")
      tradID      = rsDB(1)("tID")
      
      GetRights tradID
      
      If mainThread Then
        If sec_Trad_Admin Then
          If config_UseTrash And CLng(config_Trashbin) <> CLng(inForum) Then
            rsDB(1)("tForum") = CLng(config_Trashbin)
            rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Tråd] Raderad - Av [" & CONST_USERID & "] " & CONST_USERNAME
            rsDB(1).Update
            Con.ExeCute("UPDATE fsBB_Tradar SET tForum = " & CLng(config_Trashbin) & " WHERE tStatus_Trad = 0 AND tStatus_UnderTrad = " & tradID)
          Else
            Con.ExeCute("DELETE FROM fsBB_Tradar WHERE tStatus_Trad = 0 AND tStatus_UnderTrad = " & tradID)
            rsDB(1).Delete
          End If
          retLink = "../forum.asp?e=" & inForum
        Else
          retLink = "../trad.asp?e=" & tradID
        End If
      Else
        sec_Admin userIsAdm
        If sec_Inlagg_Admin Then
          If config_UseTrash And CLng(config_Trashbin) <> CLng(inForum) Then
            rsDB(1)("tForum") = CLng(config_Trashbin)
            rsDB(1)("tStatus_Trad") = True
            rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Svar] Raderad - Av [" & CONST_USERID & "] " & CONST_USERNAME
            rsDB(1).Update
          Else
            rsDB(1).Delete
          End If
          retLink = "../trad.asp?e=" & inUTrad
        Else
          retLink = "../trad.asp?e=" & inUTrad & "&go2=" & tradID
        End If
      End If
    End If
    
  RS_Close 1
  
  Response.Redirect(retLink)

%>