<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect("_login.asp")

  lID       = GetQ("e","123",0)   ' Titelns ID

  RS_Open 1, "SELECT * " & _
             "FROM fsBB_Tradar " & _
             "LEFT JOIN fsBB_Forum ON fID = tForum " & _
             "WHERE tID = " & CLng(lID), False

    If rsDB(1).EOF Then Response.Redirect("_err.asp")
    
    If rsDB(1)("tStatus_Trad") Then
      text_TradID = rsDB(1)("tID")
    Else
      text_TradID = rsDB(1)("tStatus_UnderTrad")
    End If
    
    GetRights text_TradID ' Hämta fram rättigheterna
    If Not sec_Trad_Visa Then Response.Redirect("_err.asp")
    
    text_ID         = CLng(rsDB(1)("tID"))
    text_Amne       = rsDB(1)("tAmne")
    text_ForumNamn  = rsDB(1)("fName")
    text_ForumID    = rsDB(1)("tForum")
    
  RS_Close 1

%>

<h3><% = sEncode(text_Amne) %></h3>
<h4>Forum: <% = sEncode(text_ForumNamn) %></h4>

<div class="popBox_Inner_Split"></div>

<form method="POST" id="popForm" target="popBox_Frame" action="/__AJAX/popbox/_action/savereport.asp">
  <label>F&ouml;rklaring till annm&auml;lan</label>
  <textarea name="xTextM" style="height: 100px;"></textarea>
  
  <input type="hidden" name="e" value=<% = lID %>>
</form>

<div class="popBox_Inner_Split"></div>

<p>Anm&auml;l bara inl&auml;gg som bryter mot v&aring;ra regler, missbruk av denna funktion kommer ge en varning.</p>

<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->