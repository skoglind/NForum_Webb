<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%
  lMedlem = CONST_USERNAME
  anvID = GetIDFromUsername(lMedlem)
  
  sQ          = Trim(MakeLegal(GetQ("q", "ABC", 255)))
  text_Forum  = GetQ("forum", "123", 0)
  If text_Forum < 0 Then text_Forum = 0
  If text_Forum = config_Trashbin Then text_Forum = 0
  
  If text_Forum > 0 Then 
    sSokIForum = "tForum = " & CLng(text_Forum)
  Else
    sSokIForum = "tForum <> " & CLng(config_Trashbin)
  End if

  If Len(sQ) > config_MinSearch Then
  
    ' #### FIX TEXT STRÄNG ####
      q = LCase(Trim(sQ))
      
      q = MakeLegal(q)
      w = Split(q, " ")
      
      For Each ww In w
        ww = Trim(ww)
        
        If Len(ww) > 2 Then
          Select Case ww
            Case Else : p = p & """*" & ww & "*"" AND "
          End Select
        End If
      Next
      
      p = Left(p, Len(p)-5)
    ' #### ^
  
    RS_Open 1, "SELECT TOP 250 tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, Rank, fIcon " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN CONTAINSTABLE(fsBB_Tradar, tTextM, '" & p & "') AS ct ON tbTrad.tID = ct.[KEY] " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "WHERE Rank > 0 AND tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND " & sSokIForum & " AND tStatus_Raderad = 0" & _
               "ORDER BY Rank DESC, tAmne ASC", False
    
      If rsDB(1).EOF Then
        any_Tradar = False
        sMess = "Inga träffar."
      Else
        any_Tradar = True
        list_Tradar = rsDB(1).GetRows
      End If
    
    RS_Close 1
  Else
    If Len(sQ) = 0 Then
      sMess = "Du har inte gjort någon sökning."
    Else
      sMess = "Inga träffar, sökordet måste bestå av fler än 3 tecken."
    End If
    any_Tradar = False
  End If
  
  RS_Open 1, "SELECT fID, fName, fSplitterBefore, fSec_Mod FROM fsBB_Forum WHERE fGroup = 0 AND (fSec_NewThread = '0' OR fSec_NewThread LIKE '%;" & SEC_TITEL & ";%') ORDER BY fNoAllView ASC, fSortNr ASC", False
    
    If rsDB(1).EOF Then
      any_Forum = False
    Else
      any_Forum = True
      list_Forum = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  sQUrl = Server.URLEncode(sQ) 
  sQ    = sEncode(sQ)
%>

<%
  ' ## Globala variabler ##
  If any_Tradar Then
    page_Title    = "[" & sQ & "] - Sök - Forum"
    page_description    = "Sök efter trådar och inlägg i forumet på N-Forum.se, Nintendo Forum. Du har just nu sökt på [" & sQ & "]"
  Else
    page_Title    = "Sök - Forum"
    page_description    = "Sök efter trådar och inlägg i forumet på N-Forum.se, Nintendo Forum. Du har inte gjort någon sökning ännu."
  End If
  
  page_Header   = "Sök i forumet"
  page_WhereAmI = "&gt; Sök i forumet "
  page_SelMenu  = "forum"
  page_Slide    = "forum"
  Remove_Distans= True
  
  page_keywords       = "sök inlägg, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1>Sök i forumet</h1>
    
      <div class="nf_msg nf_msg_full">
        <form>
           
          <select name="forum" style="width: 892px;">
            <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> Alla forum </option>
            <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
            <% If any_Forum Then %>
              <% For zx = 0 To UBound(list_Forum, 2) %>
                <% If list_Forum(2, zx) Then %><option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option><% End If %>
                <option value=<% = list_Forum(0, zx) %> style="padding: 1px 0 1px 10px;" <% If CLng(text_Forum) = CLng(list_Forum(0, zx)) Then Response.Write(" selected") %>> <% = sEncode(list_Forum(1, zx)) %> </option>
              <% Next %>
            <% End If %>
          </select> 
        
          <input style="width: 887px;" type="text" maxlength=255 name="q" value="<% = sQ %>"> 
          <input style="float: right; width: 80px; font-weight: bold;" type="submit" value="Sök">
        </form>
      </div>
    </div>
  
    <div class="nf_datablock nf_size_full">
      <div class="nf_forum nf_forum_title">
        <p style="width: 26px;">&nbsp;</p>
        <p style="width: 593px;">Foruminlägg</p>
        <p style="width: 96px;" class="nf_center">&nbsp;</p>
        <p style="width: 76px;" class="nf_center">Svar</p>
        <p style="width: 106px;" class="nf_right">Skapad</p>
      </div>
      
      <% If any_Tradar Then %>
        <% CreatePaging CONST_SET_TRADARSIDA, UBound(list_Tradar, 2) %>
        <% CreatePagingChooser %>
        
        <div class="nf_forum">
          <ul>
            <% For zx = pagingBOF To pagingEOF %>
              <%
                isTheThread = False
                If list_Tradar(4,zx) Then isTheThread = True
              %>
            
              <li class="<% If tradLastRow Then Response.Write("nf_last") %>">
                <div class="nf_icon"><img src="<% = config_GFXLocation %>icons/forum/<% = list_Tradar(8,zx) %>" alt=""></div>
                <div class="nf_amne nf_amne_short">
                  <a href="/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>"><img src="<% = config_GFXLocation %>icons/golatest.png" title="Gå till inlägget..."></a>
                  <a href="/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>"><% = sEncode(list_Tradar(1,zx)) %></a>
                  <span><% = sEncode(CutText(BBCode_Remove(list_Tradar(2,zx)),110)) %></span>
                </div>
                <div class="nf_user">&nbsp;</div>
                <div class="nf_stat"><% If isTheThread Then %><% = list_Tradar(6,zx) %><% Else %>&nbsp;<% End If %></div>
                <div class="nf_latest"><% = DatumReplace(list_Tradar(3,zx)) %></div>
              </li>
            <% Next %>
          </ul>
        </div>
        
        <div class="nf_paging nf_paging_full">
          <a href="sokforum.asp?q=<% = sQUrl %>&amp;page=<% = pagingOnPage-1 %>&amp;forum=<% = text_Forum %>">««</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="sokforum.asp?q=<% = sQUrl %>&amp;page=<% = zx %>&amp;forum=<% = text_Forum %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
            | <a href="sokforum.asp?q=<% = sQUrl %>&amp;page=<% = pagingOnPage+1 %>&amp;forum=<% = text_Forum %>">»»</a>
        </div>
        
        <% If iFilter <> 0 Then %>
          <div class="nf_forum nf_forum_bottom">
            <p style="text-align: right;"><a href="" onclick="doActionWithPrompt('_action/markasred.asp?f=<% = iFilter %>','Vill du markera alla trådar i forumet som lästa'); return false;">Markera alla trådar i forumet som lästa</a></p>
          </div>
        <% End If %>
      <% Else %>
        <div class="nf_forum nf_forum_bottom">
          <p style="text-align: center;"><% = sMess %></p>
        </div>
      <% End If %>
    </div>
  
  </div>
  
<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->