<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%
lID     = GetF("e", "123", 0)
lDO     = GetF("do", "ABC", 6)
lNewID  = GetF("newid", "123", 0)

Select Case LCase(lDO)
  Case "fuse"   : pageIs = "FUSE"
  Case "owner"  : pageIs = "OWNER"
  Case "break"  : pageIs = "BREAK"
  Case "move"   : pageIs = "MOVE"
  Case Else
    Response.Redirect("err.asp")
End Select

If CONST_LOGIN Then

  RS_Open 1, "SELECT * " & _
             "FROM fsBB_Tradar " & _
             "WHERE tID = " & CLng(lID) & " AND tStatus_Raderad = 0", True
  
    If Not rsDB(1).EOF Then
      GetRights lID ' Hämta fram rättigheterna
      If Not sec_Trad_Admin Then Response.Redirect("err.asp")
      
      Select Case pageIs
        Case "FUSE"
          If dbTradExists(lNewID) Then
            lNewForum = GetForumFromID(lNewID)
          
            rsDB(1)("tStatus_Trad")       = False
            rsDB(1)("tStatus_UnderTrad")  = lNewID
            rsDB(1)("tForum")             = lNewForum
            rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Tråd] Ihopslagen - Av [" & CONST_USERID & "] " & CONST_USERNAME
            rsDB(1).Update
            
            Con.ExeCute("UPDATE fsBB_Tradar SET tStatus_UnderTrad = " & CLng(lNewID) & ", tForum = " & CLng(lNewForum) & " WHERE tStatus_UnderTrad = " & CLng(lID))
          End if
        Case "OWNER"
          rsDB(1)("tAnv_Skapad")        = CONST_USERID
          rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Tråd] Ägarskap övertaget - Av [" & CONST_USERID & "] " & CONST_USERNAME
          rsDB(1).Update
        Case "BREAK"
          rsDB(1)("tStatus_Trad")       = True
          rsDB(1)("tStatus_UnderTrad")  = 0
          rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Svar] Utbruten till egen tråd - Av [" & CONST_USERID & "] " & CONST_USERNAME
          rsDB(1).Update
        Case "MOVE"
          If dbTradExists(lNewID) Then
            rsDB(1)("tStatus_UnderTrad")  = lNewID
            rsDB(1)("tForum")             = GetForumFromID(lNewID)
            rsDB(1)("tLogg") = rsDB(1)("tLogg") & ";" & Now & " | [Svar] Trådbyte - Av [" & CONST_USERID & "] " & CONST_USERNAME
            rsDB(1).Update
          End If
      End Select
     
    End If
    
  RS_Close 1

Else
  Response.Redirect("err.asp")
End If

%>

<script type="text/javascript">
  alert("Åtgärden utförd!");
  parent.document.getElementById("jsFrameBox").style.display = "none";
</script>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->