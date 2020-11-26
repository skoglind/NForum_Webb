<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%  
  If Not config_LockDown_Forum Then
    RS_Open 1, "SELECT TOP 100 tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, fIcon " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "WHERE tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tForum <> " & CLng(config_Trashbin) & " AND tStatus_Raderad = 0 ORDER BY tDatum_Skapad DESC", False
    
      If rsDB(1).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  
  If any_Tradar Then
    CreatePaging CONST_SET_TRADARSIDA, UBound(list_Tradar, 2)
    CreatePagingChooser
  End If
%>

<%
  ' ## Globala variabler ##
  If pagingNumOfPages > 0 Then
    page_Title    = "Nya inlägg - Sida " & pagingOnPage & " - Forumindex"
  Else
    page_Title    = "Nya inlägg - Forumindex"
  End If
  
  page_Header   = "Nya inlägg"
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> " & _
                  "&gt; Mina inlägg"
  page_SelMenu  = "forum"
  page_Slide    = "forum"
  Remove_Distans= True
  
  page_description    = "Alla de nyaste inläggen och trådarna på N-Forum.se, Nintendo Forum. Listas efter senast skapade. Du är på sida " & pagingOnPage & "."
  page_keywords       = "nya inlägg, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1>Nya inlägg</h1> 
     
      <div class="nf_forum nf_forum_title">
        <p style="width: 26px;">&nbsp;</p>
        <p style="width: 593px;">Foruminlägg</p>
        <p style="width: 96px;" class="nf_center">&nbsp;</p>
        <p style="width: 76px;" class="nf_center">Svar</p>
        <p style="width: 106px;" class="nf_right">Skapad</p>
      </div>
      
      <% If any_Tradar Then %>
        
        <div class="nf_forum">
          <ul>
            <% For zx = pagingBOF To pagingEOF %>
              <%
                isTheThread = False
                If list_Tradar(4,zx) Then isTheThread = True
              %>
            
              <li class="<% If tradLastRow Then Response.Write("nf_last") %>">
                <div class="nf_icon"><img src="<% = config_GFXLocation %>icons/forum/<% = list_Tradar(7,zx) %>" alt=""></div>
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
          <a href="nyainlagg.asp?m=<% = lMedlem %>&amp;page=<% = pagingOnPage-1 %>">««</a> |
          
          <% For Each zx In pagingPages %>
            <% If zx = "..." Then %>
              ... |
            <% Else %>
              <a href="nyainlagg.asp?m=<% = lMedlem %>&amp;page=<% = zx %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
            <% End If %>
          <% Next %>
          
          | <a href="nyainlagg.asp?m=<% = lMedlem %>&amp;page=<% = pagingOnPage+1 %>">»»</a>
        </div>
        
        <% If iFilter <> 0 Then %>
          <div class="nf_forum nf_forum_bottom">
            <p style="text-align: right;"><a href="" onclick="doActionWithPrompt('_action/markasred.asp?f=<% = iFilter %>','Vill du markera alla trådar i forumet som lästa'); return false;">Markera alla trådar i forumet som lästa</a></p>
          </div>
        <% End If %>
      <% Else %>
        <div class="nf_forum nf_forum_bottom">
          <p style="text-align: center;">Det finns inga trådar i detta forum, <a href="ny_trad.asp?f=<% = iFilter %>">bli den första att skapa en tråd</a>.</p>
        </div>
      <% End If %>
    </div>
  
  </div>
  
<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->