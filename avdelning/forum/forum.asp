<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%
  iFilter = GetQ("e", "123", 0)
  
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0
  
  nDate = DateAdd("d", -config_RemoOlasta, Now)
  
  If iFilter <> 0 Then
    sFilter = "AND tForum = " & CLng(iFilter)
    slide_status = "index"
    
    RS_Open 1, "SELECT * FROM fsBB_Forum WHERE fID = " & CLng(iFilter) & " AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%')", False
      If rsDB(1).EOF Then
        Response.Redirect("default.asp")
      Else
        sForumNamn = rsDB(1)("fName")
        sForumText = rsDB(1)("fInfo")
      End If
    Rs_Close 1
    
    sForumID   = iFilter
  Else
    sFilter = "AND fNoAllView = 0 AND tStatus_Dold = 0"
    slide_status = "allfora"
    sForumNamn = "Alla Forum"
    sForumText = "Här samlas alla vanliga forum med deras trådar och inlägg."
    sForumID   = 0
    forum_GRP  = 0
  End If
  iFilter = CLng(iFilter)
  
  ' #### Sortering av inläggen ####
  If iFilter <> 0 Then
    On Error Resume Next
      fSortering = Con.ExeCute("SELECT fSortering FROM fsBB_Forum WHERE fID = " & CLng(iFilter))(0)
    On Error Goto 0
  Else
    fSortering = 0
  End If

  Select Case fSortering
  Case 1
    sSortering = "tDatum_Skapad DESC"
  Case 2
    sSortering = "tAmne ASC"
  Case Else
    sSortering = "tDatum_Uppdaterad DESC"
  End Select
  ' #### Sortering av inläggen ####
  
  ' #### Gruppforum ####
  RS_Open 1, "SELECT * FROM fsBB_Forum WHERE fGroup = 1 AND fID = " & CLng(iFilter), False
    If Not rsDB(1).EOF Then
      sFilter = " AND tForum IN (SELECT gForum FROM fsBB_Grupper WHERE gGroup = " & CLng(rsDB(1)("fID")) & ") "
    End If
  Rs_Close 1
  '#### Gruppforum ####
  
  'WHERE oTradID = tbTrad.tID AND oAnvandare = " & CLng(lAnvID) & " AND oDatum > tbTrad.tDatum_Uppdaterad

  RS_Open 1, "SELECT tID, tAmne, A.aAnvNamn AS AaNamn, A.aID AS AaID, B.aAnvNamn AS BaNamn, B.aID AS BaID, tDatum_Uppdaterad, tInst_Klistrad, " & _ 
             "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, " & _
             "(SELECT COUNT(oID) FROM fsBB_Olast WHERE oTradID = tbTrad.tID AND oAnvandare = " & CLng(lAnvID) & " AND oDatum > tbTrad.tDatum_Uppdaterad) AS iAntalLasta, " & _
             "tStatus_Last, fIcon, " & _
             "(SELECT TOP 1 tID FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0 ORDER BY tDatum_Skapad DESC) AS iLatestThread " & _
             "FROM fsBB_Tradar AS tbTrad " & _
             "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
             "LEFT JOIN fsBB_Anv AS A ON tbTrad.tAnv_Skapad = A.aID " & _
             "LEFT JOIN fsBB_Anv AS B ON tbTrad.tAnv_Uppdaterad = B.aID " & _
             "WHERE tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tStatus_Raderad = 0 AND tStatus_Trad = 1 " & sFilter & " ORDER BY tInst_Klistrad DESC, " & sSortering, False
  
    If rsDB(1).EOF Then
      any_Tradar = False
    Else
      any_Tradar = True
      list_Tradar = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  If any_Tradar Then
    CreatePaging CONST_SET_TRADARSIDA, UBound(list_Tradar, 2)
    CreatePagingChooser
  End If
%>

<%
  ' ## Globala variabler ##
  page_Title    = sForumNamn & " - Sida " & pagingOnPage & " - Forumindex"
  page_Header   = sForumNamn
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Forumindex&quot; ...'>Forumindex</a> " & _
                  "&gt; <a href='forum.asp?e=" & sForumID & "' title='Gå till &quot;" & sForumNamn & "&quot; ...'>" & sForumNamn & "</a>"
  page_SelMenu  = "forum"
  page_Slide    = "forum"
  Remove_Distans= True
  
  page_description    = sForumNamn & ". " & sForumText & " Du är på sida " & pagingOnPage & "."
  page_keywords       = sForumNamn & ", "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">

    <div class="nf_datablock nf_size_full">
      <h1><% = sForumNamn %></h1>
     
      <div class="nf_forum nf_forum_lankar">
        <% If CONST_LOGIN Then %>
          <p><% If iFilter <> 0 Then %><a href="javascript: doActionWithPrompt('_action/markasread.asp?f=<% = iFilter %>','Vill du markera alla trådar i forumet [<% = sForumNamn %>] som lästa?');">Markera alla trådar i forumet som lästa</a> | <% End If %><a href="ny_trad.asp?f=<% = iFilter %>">Skapa en ny tråd</a></p>
        <% Else %>
          <p><a href="/avdelning/medlem/registreradig.asp" class="flash_a_red">Bli medlem GRATIS!</a> | <a href="/avdelning/medlem/loggain.asp" class="flash_a_green">Logga in</a></p>
        <% End If %>
      </div>
    
      <div class="nf_forum nf_forum_title">
        <p style="width: 26px;">&nbsp;</p>
        <p style="width: 393px;">Tråd</p>
        <p style="width: 96px;" class="nf_center">Författare</p>
        <p style="width: 76px;" class="nf_center">Svar</p>
        <p style="width: 106px;" class="nf_right">Senaste inlägg</p>
      </div>
      
      <% If any_Tradar Then %>
        
        <% For zx = pagingBOF To pagingEOF %>
          <div class="nf_forum">
            <ul>
              
              <% tradKlistrad = list_Tradar(7,zx) %>
              <% tradLastRow  = False %>
              
              <% Do Until tradKlistrad <> list_Tradar(7,zx) %>
                <%
                  If CONST_LOGIN Then
                    If list_Tradar(6,zx) > Now - config_RemoOlasta And CLng(list_Tradar(5,zx)) <> CLng(CONST_USERID) Then
                      If CLng(list_Tradar(9,zx)) = 0 Then TradIsOlast = True Else TradIsOlast = False
                    Else
                      TradIsOlast = False
                    End If
                  End If
                  
                  If CLng(zx+1) > CLng(pagingEOF) Then
                    tradLastRow = True
                  Else
                    If tradKlistrad <> list_Tradar(7,zx+1) Then tradLastRow = True
                  End If
                %>
              
                <li class="<% If tradLastRow Then Response.Write("nf_last") %>">
                  <div class="nf_icon"><img src="<% = config_GFXLocation %>icons/forum/<% = list_Tradar(11,zx) %>" alt=""></div>
                  <div class="nf_amne nf_amne_short <% If list_Tradar(10,zx) Then Response.Write(" nf_amne_locked") %>">
                    <% If TradIsOlast Then %>
                      <a href="trad.asp?e=<% = list_Tradar(0,zx) %>&amp;gl=1"><img src="<% = config_GFXLocation %>icons/golast.png" title="Gå till det senast olästa inlägget..."></a>
                    <% Else %>
                      <a href="trad.asp?e=<% = list_Tradar(0,zx) %>&amp;go2=<% = list_Tradar(12,zx) %>"><img src="<% = config_GFXLocation %>icons/golatest.png" title="Gå till det senaste inlägget..."></a>
                    <% End if %>
                    <a href="trad.asp?e=<% = list_Tradar(0,zx) %>" style="<% If TradIsOlast Then Response.Write("font-weight: bold; color: #61a02e;") %>"><% = sEncode(list_Tradar(1,zx)) %></a>
                    <%
                      AntalSvar = list_Tradar(8,zx)
                      StartNr   = RoundUp(AntalSvar, CONST_SET_INLAGGSIDA) - 4
                      SlutNr    = RoundUp(AntalSvar, CONST_SET_INLAGGSIDA)
                      If StartNr < 1 Then StartNr = 1
                    %>
                    <% If AntalSvar > CONST_SET_INLAGGSIDA Then %><span>Sida: [ <% If StartNr > 1 Then %>...<% End If %><% For xx = StartNr To SlutNr %><a href="trad.asp?e=<% = list_Tradar(0,zx) %>&amp;page=<% = xx %>"> <% = xx %></a><% Next %> ]</span><% End If %>
                  </div>
                  <div class="nf_user"><a href="/avdelning/medlem/?m=<% = sEncode(list_Tradar(2,zx)) %>"><% = sEncode(list_Tradar(2,zx)) %></a></div>
                  <div class="nf_stat"><% = list_Tradar(8,zx) %></div>
                  <div class="nf_latest">
                    <% = DatumReplace(list_Tradar(6,zx)) %><br>av 
                    <a href="/avdelning/medlem/?m=<% = sEncode(list_Tradar(4,zx)) %>"><% = sEncode(list_Tradar(4,zx)) %></a>
                  </div>
                </li>
                
                <% zx = zx + 1 %>
                <% If CLng(zx) > CLng(pagingEOF) Then Exit Do %>
              <% Loop %>
              <% zx = zx - 1 %>
              
            </ul>
          </div>
        <% Next %>
        
        <div class="nf_paging nf_paging_full">
          <a href="forum.asp?e=<% = iFilter %>&amp;page=<% = pagingOnPage-1 %>">««</a> |
          
          <% For Each zx In pagingPages %>
            <% If zx = "..." Then %>
              ... |
            <% Else %>
              <a href="forum.asp?e=<% = iFilter %>&amp;page=<% = zx %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
            <% End If %>
          <% Next %>
          
          | <a href="forum.asp?e=<% = iFilter %>&amp;page=<% = pagingOnPage+1 %>">»»</a>
        </div>
      <% Else %>
        <div class="nf_forum nf_forum_bottom">
          <p style="text-align: center;">Det finns inga trådar i detta forum, <a href="ny_trad.asp?f=<% = iFilter %>">bli den första att skapa en tråd</a>.</p>
        </div>
      <% End If %>
    </div>
  
  </div>
  
<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->