<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>
<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%

  lID   = GetQ("e","123",0)
  lStat = GetQ("s","123",0)
  retLink = "../default.asp"  

  RS_Open 1, "SELECT * FROM fsBB_Tradar WHERE tStatus_Trad = 1 AND tID = " & CLng(lID), True
    
    If Not rsDB(1).EOF Then
      tradID      = rsDB(1)("tID")
      
      GetRights tradID
      
      If sec_Trad_Admin Then
        If lStat = 1 And Not rsDB(1)("tStatus_Last")  Then
          rsDB(1)("tStatus_Last") = True
          rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Tråd] Låst - Av [" & CONST_USERID & "] " & CONST_USERNAME
          rsDB(1).Update
        ElseIf lStat = 0 And rsDB(1)("tStatus_Last") Then
          rsDB(1)("tStatus_Last") = False
          rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Tråd] Upplåst - Av [" & CONST_USERID & "] " & CONST_USERNAME
          rsDB(1).Update
        End If
        
        retLink = "../trad.asp?e=" & tradID
      End If
    End If
    
  RS_Close 1
  
  Response.Redirect(retLink)

%>