<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0

  nDate = DateAdd("d", -config_RemoOlasta, Now)

  If Not config_LockDown_Forum Then
    RS_Open 1, "SELECT fID, fName, fInfo, fGroup, fNoAllView, fDesign_BlockID, fIcon, fSec_View,  " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tDatum_Skapad <= '" & Now & "' AND tForum = fsBB_Forum.fID AND tStatus_Trad = 1 AND tStatus_Raderad = 0) AS iAntalTradar, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tDatum_Skapad <= '" & Now & "' AND tForum = fsBB_Forum.fID AND tStatus_Trad = 0 AND tStatus_Raderad = 0) AS iAntalInlagg, " & _
               
               "(SELECT COUNT(oID) FROM fsBB_Olast LEFT JOIN fsBB_Tradar ON fsBB_Tradar.tID = oTradID WHERE tDatum_Skapad <= '" & Now & "' AND fsBB_Tradar.tForum = fsBB_Forum.fID AND fsBB_Tradar.tStatus_Raderad = 0 AND oAnvandare = " & CLng(lAnvID) & " AND oDatum > fsBB_Tradar.tDatum_Uppdaterad AND fsBB_Tradar.tDatum_Uppdaterad > '" & nDate & "') AS iAntalOlasta, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tDatum_Skapad <= '" & Now & "' AND tForum = fsBB_Forum.fID AND tStatus_Raderad = 0 AND tStatus_Trad = 1 AND tDatum_Uppdaterad > '" & nDate & "') AS iAntalNyaTradar, " & _
                
               "(SELECT TOP 1 fsBB_Anv.aAnvNamn FROM fsBB_Tradar LEFT JOIN fsBB_Anv ON fsBB_Tradar.tAnv_Uppdaterad = fsBB_Anv.aID WHERE tDatum_Skapad <= '" & Now & "' AND tStatus_Trad = 1 AND tStatus_Raderad = 0 AND tForum = fsBB_Forum.fID ORDER BY tDatum_Uppdaterad DESC) AS AaNamn, " & _
               "(SELECT TOP 1 tDatum_Uppdaterad FROM fsBB_Tradar WHERE tDatum_Skapad <= '" & Now & "' AND  tStatus_Trad = 1 AND tStatus_Raderad = 0 AND tForum = fsBB_Forum.fID ORDER BY tDatum_Uppdaterad DESC) AS AaDatum " & _
                
               "FROM fsBB_Forum " & _
               "WHERE (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') " & _
               "ORDER BY fDesign_BlockID ASC, fDesign_SortNr ASC", False
    
      If rsDB(1).EOF Then
        any_Forums = False
      Else
        any_Forums = True
        list_Forums = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Forumindex"
  page_Header   = "Nintendo forum"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Forumindex&quot; ...'>Forumindex</a> "
  page_SelMenu  = "forum"
  page_Slide    = "forum"
  Remove_Distans= True
  
  page_description    = "Vårat forumindex på N-Forum.se, Nintendo Forum, där du kommer åt alla våra underforum om Nintendo."
  page_keywords       = "forumindex, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      
      <h1>Forumindex</h1>
    
      <% If Not config_LockDown_Forum Then %>
        <div class="nf_forum nf_forum_lankar">
          <% If CONST_LOGIN Then %>
            <p><a href="javascript: doActionWithPrompt('_action/markasread.asp','Vill du markera alla trådar som lästa?');">Markera alla trådar som lästa</a> | <a href="ny_trad.asp">Skapa en ny tråd</a></p>
          <% Else %>
            <p><a href="/avdelning/medlem/registreradig.asp" class="flash_a_red">Bli medlem GRATIS!</a> | <a href="/avdelning/medlem/loggain.asp" class="flash_a_green">Logga in</a></p>
          <% End If %>
        </div>
      
        <div class="nf_forum nf_forum_title">
          <p style="width: 26px;">&nbsp;</p>
          <p style="width: 613px;">Diskussionsforum</p>
          <p style="width: 76px;" class="nf_center">Trådar</p>
          <p style="width: 76px;" class="nf_center">Inlägg</p>
          <p style="width: 106px;" class="nf_right">Senaste inlägg</p>
        </div>
      <% Else %>
        <div class="nf_forum nf_forum_title">
          <p style="width: 720px;" class="nf_center">Nerstängt!<br>Forumet är tillfälligt nerstängt av N-Forum.se Admin.</p>
        </div>
      <% End If %>
    
      <% If any_Forums Then %>
        <div class="nf_forum">
          <ul>
            <li class="nf_first nf_last">
              <div class="nf_icon"><img src="<% = config_GFXLocation %>icons/forum/trad.png" alt=""></div>
              <div class="nf_amne">
                <img src="<% = config_GFXLocation %>icons/reg_nytrad.png" title="">
                <a href="forum.asp">Alla forum</a>
                <p>Visar trådar och inlägg från alla forum som inte är exkluderade.</p>
              </div>
              <div class="nf_stat">&nbsp;</div>
              <div class="nf_stat">&nbsp;</div>
              <div class="nf_latest"></div>
            </li>
          </ul>
        </div>
        
        <% For zx = 0 To Ubound(list_Forums,2) %>
          <div class="nf_forum">
            <ul>
            
              <% forumBlockID = list_Forums(5,zx) %>
              <% forumLastRow = False %>
              
              <% Do Until forumBlockID <> list_Forums(5,zx) %>
                <%
                If CONST_LOGIN Then If CLng(list_Forums(10,zx)) < CLng(list_Forums(11,zx)) Then TradIsOlast = True Else TradIsOlast = False
                
                If CLng(zx+1) > CLng(Ubound(list_Forums,2)) Then
                  forumLastRow = True
                Else
                  If forumBlockID <> list_Forums(5,zx+1) Then forumLastRow = True
                End If
                %>
                
                <li class="<% If forumLastRow Then Response.Write("nf_last") %>">
                  <% If Not list_Forums(3,zx) Then %>
                    <div class="nf_icon"><img src="<% = config_GFXLocation %>icons/forum/<% = list_Forums(6,zx) %>" alt=""></div>
                    <div class="nf_amne">
                      <% If TradIsOlast Then %>
                        <img src="<% = config_GFXLocation %>icons/nytrad.png" title="Det finns olästa inlägg i forumet...">
                      <% Else %>
                        <img src="<% = config_GFXLocation %>icons/no_nytrad.png" title="Det finns inga olästa inlägg i forumet...">
                      <% End If %>
                      <a href="forum.asp?e=<% = list_Forums(0,zx) %>" style="<% If TradIsOlast Then Response.Write("font-weight: bold; color: #61a02e;") %>"><% = list_Forums(1,zx) %></a>
                      <p><% = list_Forums(2,zx) %></p>
                    </div>
                    <div class="nf_stat"><% = list_Forums(8,zx) %></div>
                    <div class="nf_stat"><% = list_Forums(9,zx) %></div>
                    <div class="nf_latest">
                      <% If list_Forums(8,zx) > 0 Then %>
                        <% = DatumReplace(list_Forums(13,zx)) %><br>av 
                        <% lNamn = "" : lNamn = list_Forums(12,zx) %>
                        <a href="/avdelning/medlem/?m=<% = sEncode(lNamn) %>"><% = sEncode(lNamn) %></a>
                      <% Else %>
                        Inga inlägg
                      <% End If %>
                    </div>
                  <% Else %>
                    <div class="nf_onerow"><a href="forum.asp?e=<% = list_Forums(0,zx) %>" style="<% If TradIsOlast Then Response.Write("font-weight: bold; color: #61a02e;") %>"><% = list_Forums(1,zx) %></a></div>
                  <% End If %>
                </li>
                
                <% zx = zx + 1 %>
                <% If CLng(zx) > CLng(Ubound(list_Forums,2)) Then Exit Do %>
              <% Loop %>
              
              <% zx = zx - 1 %>
              
            </ul>
          </div>
        <% Next %>
        
      <% End If %>
      
      <% If Not config_LockDown_Forum Then %>
        <div class="nf_forum nf_forum_bottom">
          <%
          totInlagg = con.ExeCute("SELECT COUNT(tID) FROM fsBB_Tradar WHERE tDatum_Skapad <= '" & Now & "' AND tStatus_Trad = 0 AND tStatus_Raderad = 0")(0)
          totTrad   = con.ExeCute("SELECT COUNT(tID) FROM fsBB_Tradar WHERE tDatum_Skapad <= '" & Now & "' AND tStatus_Trad = 1 AND tStatus_Raderad = 0")(0)
          totMedlem = con.ExeCute("SELECT COUNT(aID) FROM fsBB_Anv WHERE aBlockadTill < '" & Date & "' AND aAktiverad = 1")(0)
          %>
          <p>Totalt antal inlägg <strong><% = totInlagg %></strong> | Totalt antal trådar <strong><% = totTrad %></strong> <% If CONST_LOGIN Then %>| Totalt antal medlemmar <strong><% = totMedlem %></strong></p><% End If %>
        </div>
      <% End If %>
      
    </div>
  
  </div>
  
<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->