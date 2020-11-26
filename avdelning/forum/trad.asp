<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%
  iID       = GetQ("e","123",0)
  postID    = GetQ("go2","123",0)
  gotoLast  = GetQ("gl","123",0) 
  lLastID   = 0
  bGoAway   = False

  ' #### Användaren har efterfrågat en speciell tråd och ska givetvist skickas dit
  If postID > 0 And postID <> iID Then
    RS_Open 1, "SELECT * FROM fsBB_Tradar WHERE tStatus_UnderTrad = " & CLng(iID) & " AND tStatus_Raderad = 0 ORDER BY tDatum_Skapad ASC", True
  
      If Not rsDB(1).EOF Then      
        iPerSida  = CONST_SET_INLAGGSIDA
        rsDB(1).Filter = "tID = " & CLng(postID) & ""
        iCurPos   = rsDB(1).AbsolutePosition
      
        nUseSida = RoundUp(iCurPos, iPerSida)
        
        Response.Redirect("trad.asp?e=" & iID & "&page=" & nUseSida & "#ID" & postID)
      End If
      
    RS_Close 1
  End If
  ' #### Jupp ^^
  
  ' #### Hämta senast lästa inläggid och skicka användaren ett snäpp längre
  
  If gotoLast = 1 AND CONST_LOGIN Then
    dCount = Con.ExeCute("SELECT COUNT(oDatum) FROM fsBB_Olast WHERE oAnvandare = " & CLng(CONST_USERID) & " AND oTradID = " & CLng(iID))(0)
    If dCount > 0 Then
      dOlast = Con.ExeCute("SELECT oDatum FROM fsBB_Olast WHERE oAnvandare = " & CLng(CONST_USERID) & " AND oTradID = " & CLng(iID))(0)
      If Not IsDate(dOlast) Then dOlast = Now
      
      dNoPlus = False
    Else
      dCountI = Con.ExeCute("SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = " & CLng(iID))(0)
      If dCountI > 0 Then dOlast = Con.ExeCute("SELECT TOP 1 tDatum_Skapad FROM fsBB_Tradar WHERE tStatus_UnderTrad = " & CLng(iID) & " ORDER BY tDatum_Skapad ASC")(0)
      If Not IsDate(dOlast) Then dOlast = Now
      
      dNoPlus = True
    End If
    
    RS_Open 1, "SELECT tID, tDatum_Skapad " & _
               "FROM fsBB_Tradar " & _
               "WHERE tDatum_Skapad >= '" & CDate(dOlast) & "' AND tStatus_UnderTrad = " & CLng(iID) & " ORDER BY tDatum_Skapad ASC", True
               
      If Not rsDB(1).EOF Then
        If rsDB(1).RecordCount > 1 AND dNoPlus = False Then rsDB(1).MoveNext
        lLastID = rsDB(1)("tID")
        
        bGoAway = True
      End If
    RS_Close 1
    
    If bGoAway Then
      Response.Redirect("trad.asp?e=" & iID & "&go2=" & lLastID)
    End If
  End If
  
  ' #### Jupp ^^
  
  RS_Open 1, "SELECT *, fsBB_Anv.*, B.aAnvNamn AS BaNamn, " & _
             "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tAnv_Skapad = fsBB_Anv.aID AND tStatus_Raderad = 0) AS iAntalInlagg, " & _
             "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_Raderad = 0) AS iTotInlagg " & _
             "FROM fsBB_Tradar " & _
             "LEFT JOIN fsBB_Anv ON fsBB_Tradar.tAnv_Skapad = aID " & _ 
             "LEFT JOIN fsBB_Anv AS B ON fsBB_Tradar.tAnv_Andrad = B.aID " & _ 
             "LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID " & _
             "LEFT JOIN fsBB_Forum ON tForum = fsBB_Forum.fID " & _
             "WHERE tDatum_Skapad <= '" & Now & "' AND tID = " & CLng(iID) & " AND tStatus_Raderad = 0 AND tStatus_Trad = 1", False
  
    If Not rsDB(1).EOF Then
      GetRights iID ' Hämta fram rättigheterna
      If Not sec_Trad_Visa Then Response.Redirect("default.asp")
    
      text_ID         = CLng(rsDB(1)("tID"))
    
      text_Amne       = sEncode(rsDB(1)("tAmne"))
      text_Data       = BBCode(sEncode(rsDB(1)("tTextM")), rsDB(1)("tInst_Smilies"))
      text_Signatur   = sEncode(rsDB(1)("aSignatur"))
      text_AnvNamn    = sEncode(rsDB(1)("aAnvNamn"))
      text_EgenTitel  = sEncode(rsDB(1)("aEgenTitel"))
      text_AnvID      = sEncode(rsDB(1)("aID"))
      text_Datum      = rsDB(1)("tDatum_Skapad")
      text_DatumCh    = rsDB(1)("tDatum_Andrad")
      text_AnvNamnCh  = rsDB(1)("BaNamn")
      
      text_Logg       = Replace(rsDB(1)("tLogg") & " ", ";", "<br>")
      
      text_ForumNamn  = rsDB(1)("fName")
      text_ForumID    = rsDB(1)("tForum")
      text_ForumIcon  = rsDB(1)("fIcon")
      
      text_Last       = rsDB(1)("tStatus_Last")
      
      text_bAvatar    = rsDB(1)("aAvatar")
      
      text_Plats      = rsDB(1)("aPlats")
      text_aTimeStamp = rsDB(1)("aTimeStamp")
      text_AktiveraPM = rsDB(1)("aAktiveraPM")
      
      If text_Last Then
        sText   = "Lås upp tråden"
        sLankT  = "Lås upp"
        sLank   = "_action/lock.asp?e=" & text_ID & "&amp;s=0"
      Else
        sText   = "Lås tråden"
        sLankT  = "Lås"
        sLank   = "_action/lock.asp?e=" & text_ID & "&amp;s=1"
      End If

    Else
      Response.Redirect("default.asp")
    End If
  
  RS_Close 1

  RS_Open 1, "SELECT tID, tAmne, tTextM, B.aAnvNamn AS BaNamn, fsBB_Anv.aAnvNamn, fsBB_Anv.aSignatur, tDatum_Skapad, fsBB_Anv.aID, " & _
             "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tAnv_Skapad = fsBB_Anv.aID AND tStatus_Raderad = 0) AS iAntalInlagg, " & _
             "tInst_Smilies, ttAdmin, tDatum_Andrad, tLogg, fsBB_Anv.aAvatar, fsBB_Anv.aPlats, fsBB_Anv.aTimeStamp, fsBB_Anv.aAktiveraPM, fsBB_Anv.aEgenTitel " & _
             "FROM fsBB_Tradar " & _ 
             "LEFT JOIN fsBB_Anv ON fsBB_Tradar.tAnv_Skapad = aID " & _ 
             "LEFT JOIN fsBB_Anv AS B ON fsBB_Tradar.tAnv_Andrad = B.aID " & _ 
             "LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID " & _
             "WHERE tDatum_Skapad <= '" & Now & "' AND tStatus_UnderTrad = " & CLng(iID) & " AND tStatus_Raderad = 0 AND tStatus_Trad = 0 ORDER BY tDatum_Skapad ASC", False
  
    If rsDB(1).EOF Then
      any_Inlagg = False
    Else
      any_Inlagg = True
      list_Inlagg = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  If CONST_LOGIN Then
    Con.ExeCute("DELETE FROM fsBB_Olast WHERE oDatum < '" & Now - config_RemoOlasta & "' OR (oTradID = " & CLng(iID) & " AND oAnvandare = " & CLng(CONST_USERID) & ")")
    Con.ExeCute("INSERT INTO fsBB_Olast (oTradID, oDatum, oAnvandare) VALUES(" & CLng(iID) & ",'" & DateAdd("n", 1,Now) & "'," & CLng(CONST_USERID) & ")")
  End If
  
  If any_Inlagg Then
    CreatePaging CONST_SET_INLAGGSIDA, UBound(list_Inlagg, 2)
    CreatePagingChooser
  End If
%>

<%
  ' ## Globala variabler ##
  If pagingNumOfPages > 0 Then
    page_Title    = text_Amne & " - ID:" & iID & " - Sida " & pagingOnPage & " - " & text_ForumNamn & " - Forumindex"
  Else
    page_Title    = text_Amne & " - " & text_ForumNamn & " - Forumindex"
  End If
  page_Header   = text_Amne
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Forumindex&quot; ...'>Forumindex</a> " & _
                  "&gt; <a href='forum.asp?e=" & text_ForumID & "' title='Gå till &quot;" & text_ForumNamn & "&quot; ...'>" & text_ForumNamn & "</a> " & _
                  "&gt; <a href='trad.asp?e=" & iID & "' title='Gå till &quot;" & text_Amne & "&quot; ...'>" & text_Amne & "</a>"
  page_SelMenu  = "forum"
  page_Slide    = "forum"
  Remove_Distans= True
  
  iFilter       = text_ForumID
  
  page_description    = "Ämnet (" & text_Amne & ", " & iID & ") diskuteras i forumet " &  text_ForumNamn & ". Du är på sida " & pagingOnPage & "."
  page_keywords       = text_ForumNamn & ", "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_twothird"><h1><% = text_Amne %></h1></div>
    <div class="nf_datablock nf_size_onethird" style="text-align: right;"><h1>#<% = iID %></h1></div>
  
    <div class="nf_datablock nf_size_full">
      <h2><a href="default.asp">Forumindex</a> / <a href="forum.asp?e=<% = text_ForumID %>"><% = text_ForumNamn %></a></h2>
    </div>
    
    <div class="nf_datablock nf_size_full">
      <div class="nf_forum nf_forum_lankar">
        <% If CONST_LOGIN Then %>
          <p><a href="ny_trad.asp?a=<% = text_ID %>&amp;url=<% = onURL %>" title="Svara på tråden">Skriv ett inlägg i tråden</a> | <a href="ny_trad.asp?f=<% = iFilter %>">Skapa en ny tråd</a></p>
        <% Else %>
          <p><a href="/avdelning/medlem/registreradig.asp" class="flash_a_red">Bli medlem GRATIS!</a> | <a href="/avdelning/medlem/loggain.asp" class="flash_a_green">Logga in</a></p>
        <% End If %>
      </div>
    </div>
    
    <div class="nf_datablock nf_size_full">
      
      <div class="nf_forum">
      
        <div class="nf_trad">
          <div class="nf_trad_content">
            <p class="nf_trad_title"><% = text_Amne %></p>
            <p><% = text_Data %></p>
            <div class="nf_trad_logg" id="perm<% = text_ID %>"><p><strong>http://<% = page_SubDomain %>n-forum.se/avdelning/forum/trad.asp?e=<% = text_ID %></strong></p></div>
            <% If sec_Trad_Admin Then %>
              <div class="nf_trad_logg" id="log<% = text_ID %>">
                <p><strong>Loggade händelser för denna post</strong></p>
                <p><% = text_Logg %></p>
              </div>
            <% End If %>
            <% If DateAdd("n", -10, text_DatumCh) > text_Datum Then %><p class="nf_trad_chg">Senast ändrad <% = DatumReplace(text_DatumCh) %> av <% = text_AnvNamnCh %></p><% End If %>
            <% If Len(text_Signatur) > 0 AND CONST_SET_SIGN Then %><p class="nf_trad_sign"><% = text_Signatur %></p><% End If %>
          </div>
          
          <div class="nf_trad_data">
            <% If text_bAvatar AND CONST_SET_AVATAR Then %>
              <div class="nf_trad_avatar"><img src="<% = config_Avatar %>u<% = Right("000000" & text_AnvID, 6) %>.jpg" alt="Avatar" title=""></div>
            <% End If %>
            <div class="nf_trad_username"><a href="/avdelning/medlem/?m=<% = text_AnvNamn %>"><% = text_AnvNamn %></a></div>
            <% If Len(text_EgenTitel) > 1 Then %><div class="nf_trad_egentitel"><% = text_EgenTitel %></div><% End If %>
            <div class="nf_trad_otherinfo">
              <% If Len(text_Plats) > 0 Then %><p><strong>Plats:</strong> <% = sEncode(text_Plats) %> </p><% End If %>
              <p style="margin-bottom: 6px;"><strong>Status:</strong> <% If text_aTimeStamp > DateAdd("n", -5, Now) Then %><span style='color: #0A0; font-weight: bold;'>Online</span><% Else %><span style='color: #A00;'>Offline</span><% End If %> </p> 
              <% If text_AktiveraPM And CONST_LOGIN Then %><p> <a href="/avdelning/medlem/skrivpm.asp?m=<% = text_AnvNamn %>">» Skicka PM</a> </p><% End If %>
              <p> <a href="/avdelning/medlem/minaspel.asp?m=<% = text_AnvNamn %>">» Se spellista</a> </p> 
            </div>
          </div>
        
          <div class="nf_trad_control">
            <div class="nf_trad_date"><% = DatumReplace(text_Datum) %></div>
            <div class="nf_trad_admin"><% If sec_Trad_Admin Then %><a href="<% = sLank %>" title="<% = sText %>"><% = sLankT %></a> | <% End If %> <% If sec_Trad_Hantera Then %><a href="ny_trad.asp?e=<% = text_ID %>&amp;url=<% = onURL %>" title="Redigera tråden">Ändra</a><% End If %> <% If sec_Trad_Admin Then %> | <a onclick="showSwap('log<% = text_ID %>'); return false;" title="Visa/Dölj Logg" style="cursor: pointer;">Logg</a> | <a href="" onclick="doActionWithPrompt('_action/delete.asp?e=<% = text_ID %>','Vill du radera tråden?'); return false;" title="Radera tråden" class="red">X</a><% End If %>&nbsp;</div>
            <div class="nf_trad_user">&nbsp;<a onclick="showSwap('perm<% = text_ID %>'); return false;" title="Visa/Dölj Permalänk" style="cursor: pointer;">Permalänk</a>  <% If sec_Inlagg_Skapa Then %>| <a href="ny_trad.asp?a=<% = text_ID %>&amp;url=<% = onURL %>" title="Svara på tråden">Svara</a> | <a href="ny_trad.asp?c=<% = text_ID %>&amp;url=<% = onURL %>" title="Citera tråden">Citera</a> | <% End If %> <% If CONST_LOGIN Then %><a href="javascript: OpenReportPost(<% = text_ID %>);" title="Anmäl tråden" class="red">Anmäl</a><% End If %></div>
          </div>
        </div>
      
      </div>
      
      <% If any_Inlagg Then %>
      
        <div class="nf_paging nf_paging_full">
          <a href="trad.asp?e=<% = iID %>&amp;page=<% = pagingOnPage-1 %>">««</a> |
              
          <% For Each zx In pagingPages %>
            <% If zx = "..." Then %>
              ... |
            <% Else %>
              <a href="trad.asp?e=<% = iID %>&amp;page=<% = zx %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
            <% End If %>
          <% Next %>
          
          | <a href="trad.asp?e=<% = iID %>&amp;page=<% = pagingOnPage+1 %>">»»</a>
        </div>
        
        <% For zx = pagingBOF To pagingEOF %>
          <% If zx > UBound(list_Inlagg, 2) Then Exit For %>
          <%
            bMeHantera = sec_Hantera(list_Inlagg(10,zx),list_Inlagg(7,zx))
            bMeAdmin   = sec_Admin(list_Inlagg(10,zx))
          %>
          
          <div class="nf_forum">
           
            <div class="nf_trad">
              <div class="nf_trad_content">
                <p class="nf_trad_title"><% = sEncode(list_Inlagg(1,zx)) %><a name="ID<% = list_Inlagg(0,zx) %>"> </a></p>
                <p><% = BBCode(sEncode(list_Inlagg(2,zx)), list_Inlagg(9,zx)) %></p>
                <div class="nf_trad_logg" id="perm<% = list_Inlagg(0,zx) %>"><p><strong>http://<% = page_SubDomain %>n-forum.se/avdelning/forum/trad.asp?e=<% = text_ID %>&go2=<% = list_Inlagg(0,zx) %></strong></p></div>
                <% If sec_Trad_Admin Then %>
                  <div class="nf_trad_logg" id="log<% = list_Inlagg(0,zx) %>">
                    <p><strong>Loggade händelser för denna post</strong></p>
                    <p><% = Replace(list_Inlagg(12,zx) & " ", ";", "<br>") %></p>
                  </div>
                <% End If %>
                <% If DateAdd("n", -10, list_Inlagg(11,zx)) > list_Inlagg(6,zx) Then %><p class="nf_trad_chg">Senast ändrad <% = DatumReplace(list_Inlagg(11,zx)) %> av <% = list_Inlagg(3,zx) %></p><% End If %>
                <% If Len(list_Inlagg(5,zx)) > 0 AND CONST_SET_SIGN Then %><p class="nf_trad_sign"><% = sEncode(list_Inlagg(5,zx)) %></p><% End If %>
              </div>
              
              <div class="nf_trad_data">
                <% If list_Inlagg(13,zx) AND CONST_SET_AVATAR Then %>
                  <div class="nf_trad_avatar"><img src="<% = config_Avatar %>u<% = Right("000000" & list_Inlagg(7,zx), 6) %>.jpg" alt="Avatar" title=""></div>
                <% End If %>
                <div class="nf_trad_username"><a href="/avdelning/medlem/?m=<% = list_Inlagg(4,zx)  %>"><% = list_Inlagg(4,zx) %></a></div>
                <% If Len(list_Inlagg(17,zx)) > 1 Then %><div class="nf_trad_egentitel"><% = list_Inlagg(17,zx) %></div><% End If %>
                <div class="nf_trad_otherinfo">
                  <% If Len(list_Inlagg(14,zx)) > 0 Then %><p><strong>Plats:</strong> <% = sEncode(list_Inlagg(14,zx)) %> </p><% End If %>
                  <p style="margin-bottom: 6px;"><strong>Status:</strong> <% If list_Inlagg(15,zx) > DateAdd("n", -5, Now) Then %><span style='color: #0A0; font-weight: bold;'>Online</span><% Else %><span style='color: #A00;'>Offline</span><% End If %> </p> 
                  <% If list_Inlagg(16,zx) And CONST_LOGIN Then %><p> <a href="/avdelning/medlem/skrivpm.asp?m=<% = list_Inlagg(4,zx) %>">» Skicka PM</a> </p><% End If %>
                  <p> <a href="/avdelning/medlem/minaspel.asp?m=<% = list_Inlagg(4,zx) %>">» Se spellista</a> </p> 
                </div>
              </div>
            
              <div class="nf_trad_control">
                <div class="nf_trad_date"><% = DatumReplace(list_Inlagg(6,zx)) %></div>
                <div class="nf_trad_admin"><% If bMeHantera Then %><a href="ny_trad.asp?e=<% = list_Inlagg(0,zx) %>&amp;url=<% = onURL %>" title="Redigera inlägget">Ändra</a><% End If %> <% If bMeAdmin Then %> | <a onclick="showSwap('log<% = list_Inlagg(0,zx) %>'); return false;" title="Visa/Dölj Logg" style="cursor: pointer;">Logg</a> | <a href="" onclick="doActionWithPrompt('_action/delete.asp?e=<% = list_Inlagg(0,zx) %>','Vill du radera inlägget?'); return false;" title="Radera inlägget" class="red">X</a><% End If %>&nbsp;</div>
                <div class="nf_trad_user">&nbsp;<a onclick="showSwap('perm<% = list_Inlagg(0,zx) %>'); return false;" title="Visa/Dölj Permalänk" style="cursor: pointer;">Permalänk</a> <% If sec_Inlagg_Skapa Then %>| <a href="ny_trad.asp?a=<% = text_ID %>&amp;url=<% = onURL %>" title="Svara på inlägget">Svara</a> | <a href="ny_trad.asp?c=<% = list_Inlagg(0,zx) %>&amp;url=<% = onURL %>" title="Citera inlägget">Citera</a> | <% End If %> <% If CONST_LOGIN Then %><a href="javascript: OpenReportPost(<% = list_Inlagg(0,zx) %>);" title="Anmäl inlägget" class="red">Anmäl</a><% End If %></div>
              </div>
            </div>
          
          </div>
        <% Next %>
        
        
        <div class="nf_paging nf_paging_full">
          <a href="trad.asp?e=<% = iID %>&amp;page=<% = pagingOnPage-1 %>">««</a> |
              
          <% For Each zx In pagingPages %>
            <% If zx = "..." Then %>
              ... |
            <% Else %>
              <a href="trad.asp?e=<% = iID %>&amp;page=<% = zx %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
            <% End If %>
          <% Next %>
          
          | <a href="trad.asp?e=<% = iID %>&amp;page=<% = pagingOnPage+1 %>">»»</a>
        </div>
        
      <% End If %>
      
      <% If CONST_LOGIN And sec_Inlagg_Skapa Then %>
      
        <div class="nf_forum" id="edit_preview">
          <form method="POST" action="_action/post.asp" id="frmQuickPost">
            <p class="nf_p_full"><strong>Snabbsvar</strong><a name="EDIT"> </a></p>
            <div class="nf_post_msg"><textarea name="fMsg" id="fMsg" maxlength="20000" onkeyup="if(this.value==''){document.getElementById('btPreview').disabled=true;}else{document.getElementById('btPreview').disabled=false;}closePreview('fMsg',false);"></textarea></div>
            <div class="nf_post_btn">
              <input type="submit" style="font-weight: bold;" value="Posta">
              <input type="button" id="btPreview" disabled value="Förhandsgranska" onclick="doPreview('fMsg','YES','YES');">
            </div>
            <input type="hidden" name="fAmne" value="<% = "Sv: " & text_Amne %>">
            <input type="hidden" name="tradID_Svar" value=<% = iID %>>
            <input type="hidden" name="fAutoUrl" value="YES">
            <input type="hidden" name="fAutoSmil" value="YES">
          </form>
        </div>
        
        <div class="nf_msg_full nf_green" id="warn_preview" style="display: none;">
          <p><img src="<% = config_GFXLocation %>preview_arrow.png" style="float: left; margin-right: 8px;"><strong>Din förhandsgranskning, klicka på [Ändra] för att fortsätta redigera din post.</strong></p>
        </div>
        
        <div class="nf_forum" id="post_preview" style="display: none;">
           
          <div class="nf_trad">
            <div class="nf_trad_content">
              <p class="nf_trad_title"><% = "Sv: " & text_Amne %><a name="PREVIEW"> </a></p>
              <p id="post_preview_text"></p>
            </div>
            
            <div class="nf_trad_data">
              <div class="nf_post_btn">
                <input type="button" style="font-weight: bold;" value="Posta" onclick="document.getElementById('frmQuickPost').submit();">
                <input type="button" value="Ändra" onclick="closePreview('fMsg',true);">
              </div>
            </div>
          </div>
        
        </div>
      <% ElseIf Not CONST_LOGIN Then %>
        <div class="nf_msg_full nf_green">
          <p style="text-align: center;">Du måste <strong><a href="/avdelning/medlem/loggain.asp">logga in</a></strong> för att kunna skriva i forumet.</p>
          <p style="text-align: center;">Om du inte redan har en användare kan du <strong><a href="/avdelning/medlem/registreradig.asp">bli medlem</a> GRATIS!</strong>.</p>
        </div>
      <% End If %>
    
    </div>
    
  </div>
  
<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->