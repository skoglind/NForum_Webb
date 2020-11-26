<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  lMedlem = GetQ("m","ABC",50)
  If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME

  If Not dbUserExists(lMedlem) Then Response.Redirect("/")
  anvID = GetIDFromUsername(lMedlem)
  
  filter_sort   = CLng(GetQ("sort","123",0))
  filter_konsol = CLng(GetQ("k","123",0))
  filter_list   = LCase(GetQ("list","ABC", 15))
  
  Select Case filter_list
    Case "konsol"     : sList = "KONSOL"
    Case "tillbehor"  : sList = "TILLBEHOR"
    Case Else         : sList = "SPEL"
  End Select
  
  Select Case sList
    Case "SPEL"
      ' #### SPELSAMLING
      If filter_konsol > 11 Or filter_konsol < 0 Then filter_konsol = 0
      Select Case filter_konsol
        Case 1,2,3,4,5,6,7,8,9,10,11  : nKonsol = " AND sKonsol = " & CLng(filter_konsol) & " "
        Case Else                     : nKonsol = " "
      End Select
      
      If filter_sort > 3 Or filter_sort < 0 Then filter_sort = 0
      Select Case filter_sort
        Case 1    : nSort = "tRegion ASC, tTitel ASC"
        Case 2    : nSort = "tRelease ASC, tTitel ASC"
        Case 3    : nSort = "sKonsol ASC, tTitel ASC"
        Case Else : nSort = "tTitel ASC, tRegion ASC"
      End Select
      
      RS_Open 1, "SELECT biID, biTitelID, biBox, biManual, biMedia, biExtra, biInPris, tTitel, tRegion, sKonsol, tExtra " & _ 
                 "FROM cms_Bind_Anv_Spel " & _
                 "LEFT JOIN cms_Speltitlar ON cms_Bind_Anv_Spel.biTitelID = cms_Speltitlar.tID " & _
                 "LEFT JOIN cms_Spel ON cms_Speltitlar.tSpelID = cms_Spel.sID " & _
                 "WHERE sSynlig = 1 AND biAnv = " & CLng(anvID) & nKonsol & _
                 "ORDER BY tTitel ASC", False
                   
        If rsDB(1).EOF Then
          any_Samling = False
        Else
          any_Samling = True
          list_Samling = rsDB(1).GetRows()
        End If
      
      RS_Close 1
      
      hasExtra = True
      sBild = "spel"
      sEditWr = "game"
      ' #### #### ####
    Case "KONSOL"
      ' #### KONSOLSAMLING
      If filter_sort > 3 Or filter_sort < 0 Then filter_sort = 0
      Select Case filter_sort
        Case 1    : nSort = "tRegion ASC, tTitel ASC"
        Case 2    : nSort = "tRelease ASC, tTitel ASC"
        Case 3    : nSort = "kKonsol ASC, tTitel ASC"
        Case Else : nSort = "tTitel ASC, tRegion ASC"
      End Select
      
      RS_Open 1, "SELECT biID, biTitelID, biBox, biManual, biKonsol, biExtra, biInPris, tTitel, tRegion, kKonsol " & _ 
                 "FROM cms_Bind_Anv_Konsol " & _
                 "LEFT JOIN cms_Konsoltitlar ON cms_Bind_Anv_Konsol.biTitelID = cms_Konsoltitlar.tID " & _
                 "LEFT JOIN cms_Konsol ON cms_Konsoltitlar.tKonsolID = cms_Konsol.kID " & _
                 "WHERE kSynlig = 1 AND biAnv = " & CLng(anvID) & nKonsol & _
                 "ORDER BY tTitel ASC", False
                   
        If rsDB(1).EOF Then
          any_Samling = False
        Else
          any_Samling = True
          list_Samling = rsDB(1).GetRows()
        End If
      
      RS_Close 1
      sBild = "konsol"
      sEditWr = "console"
      ' #### #### ####
    Case "TILLBEHOR"
      ' #### TILLBEHÖRSAMLING
      If filter_sort > 3 Or filter_sort < 0 Then filter_sort = 0
      Select Case filter_sort
        Case 1    : nSort = "tRegion ASC, tTitel ASC"
        Case 2    : nSort = "tRelease ASC, tTitel ASC"
        Case 3    : nSort = "iKonsol ASC, tTitel ASC"
        Case Else : nSort = "tTitel ASC, tRegion ASC"
      End Select
      
      RS_Open 1, "SELECT biID, biTitelID, biBox, biManual, biTillbehor, biExtra, biInPris, tTitel, tRegion, iKonsol " & _ 
                 "FROM cms_Bind_Anv_Tillbehor " & _
                 "LEFT JOIN cms_Tillbehortitlar ON cms_Bind_Anv_Tillbehor.biTitelID = cms_Tillbehortitlar.tID " & _
                 "LEFT JOIN cms_Tillbehor ON cms_Tillbehortitlar.tTillbehorID = cms_Tillbehor.iID " & _
                 "WHERE iSynlig = 1 AND biAnv = " & CLng(anvID) & nKonsol & _
                 "ORDER BY tTitel ASC", False
                   
        If rsDB(1).EOF Then
          any_Samling = False
        Else
          any_Samling = True
          list_Samling = rsDB(1).GetRows()
        End If
      
      RS_Close 1
      sBild = "tillbehor"
      sEditWr = "addon"
      ' #### #### ####
  End Select
  
  text_AntalSpel      = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Spel WHERE biAnv = " & CLng(anvID))(0)
  text_AntalKonsoler  = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Konsol WHERE biAnv = " & CLng(anvID))(0)
  text_AntalTillbehor = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Tillbehor WHERE biAnv = " & CLng(anvID))(0)
  
  If CLng(anvID) <> CONST_USERID Then canEdit = False Else canEdit = True
  filterAll = "&amp;m=" & lMedlem & "&amp;k=" & filter_konsol & "&amp;sort=" & filter_sort & "&amp;list=" & LCase(sList)
%>
  
<%

  ' ## Globala variabler ##
  iPageData     = Request.QueryString("page")
  If Not IsNumeric(iPageData) Or Len(iPageData) = 0 Then iPageData = 1
  
  page_Title    = lMedlem & " - Spelsamling - [Sida " & iPageData & "] - Medlem"
  page_Header   = lMedlem & "s spelsamling"
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> " & _
                  "&gt; Spelsamling"
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <% If CONST_LOGIN And CLng(anvID) = CLng(CONST_USERID) Then %>
    <!--#INCLUDE FILE="__menu_u.asp"-->
  <% Else %>
    <!--#INCLUDE FILE="__menu_other.asp"-->
  <% End If %>
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">

    <div class="nf_datablock nf_size_full">
      <h1><% = lMedlem %>s spellista</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <% If CLng(anvID) = CLng(CONST_USERID) Then %><div class="nf_msg"><p><strong>Permalänk:</strong> http://<% = page_NForum %>/q/samling/?m=<% = sEncode(lMedlem) %></p></div><% End If %>
      
      <% If hasExtra And lstKonsol(filter_konsol) <> lstKonsol(0) Then %><h2><% = lstKonsol(filter_konsol) %></h2><% End If %>
      
      <% ' #### SAMLINGEN #### %>
      <div class="nf_msg nf_green">
        <ul class="nf_rowlist" id="titleListed_List" style="<% If Not any_Samling Then Response.Write("display: none;") %>">
          <% If any_Samling Then %>
            <% For zx = 0 To UBound(list_Samling, 2) %>
              <li id="titleListed_Row_<% = list_Samling(0, zx) %>">
              
                <% titleTT = sEncode(CutText(list_Samling(7, zx), 65)) & "</a>" %>
                <% If hasExtra Then %><% If Len(list_Samling(10, zx)) > 0 Then titleTT = titleTT & "<span> - " & sEncode(list_Samling(10, zx)) & "</span>" %><% End If %>
                <% If list_Samling(2, zx) = True Then cBox = "blank" Else cBox = "" %>
                <% If list_Samling(4, zx) = True Then cMedia = "blank" Else cMedia = "" %>
                <% If list_Samling(3, zx) = True Then cManual = "blank" Else cManual = "" %>
                <% If list_Samling(5, zx) = True Then cExtra = "blank" Else cExtra = "" %>
              
                <img style="width: 16px; height: 16px;" src="<% = config_GFXLocation %>icons/konsol/<% = list_Samling(9, zx) %>.png" alt="" title="">
                <img src="<% = config_GFXLocation %>icons/flags/<% = list_Samling(8, zx) %>.png" alt="" title="">
                <a href="/avdelning/<% = sBild %>/<% = sBild %>_visa_info.asp?e=<% = list_Samling(1, zx) %>" title="<% = sEncode(list_Samling(7, zx)) %>"><% = titleTT %></a>
                <% If canEdit Then %>
                  <span style="float: right;">
                    <img src="<% = config_GFXLocation %>icons/redigera.gif" alt="R" title="Redigera" title="" onclick="OpenCollection('<% = sEditWr %>',<% = list_Samling(1, zx) %>,<% = list_Samling(0, zx) %>,'edit');">
                    <img src="<% = config_GFXLocation %>icons/radera.gif" alt="X" title="Radera" onclick="DeleteCollection('<% = sEditWr %>',<% = list_Samling(0, zx) %>);">
                  </span>
                <% End If %>
                <div class="nf_collectionbar" style="background-image: url('<% = config_GFXLocation %>icons/samling/samling_alla_<% = sBild %>.png');">
                  <img alt="" title="Box" src="<% = config_GFXLocation %>icons/samling/no<% = cBox %>.png">
                  <img alt="" title="Media" src="<% = config_GFXLocation %>icons/samling/no<% = cMedia %>.png">
                  <img alt="" title="Manual" src="<% = config_GFXLocation %>icons/samling/no<% = cManual %>.png"> 
                  <img alt="" title="Extra" src="<% = config_GFXLocation %>icons/samling/no<% = cExtra %>.png"> 
                </div>
              </li>
            <% Next %>
          <% End If %>
        </ul>   
        <p class="nf_pretend_rowlist" id="titleListed_Mess" style="<% If any_Samling Then Response.Write("display: none;") %>">Det finns inga titlar att visa.</p>
      </div>
      
      <div id="titleListed_Clone" style="display: none;">
        <img src="<% = config_GFXLocation %>icons/konsol/XXXX_KONSOL.png" alt="" title="">
        <img src="<% = config_GFXLocation %>icons/flags/XXXX_REGION.png" alt="" title="">
        <a href="/avdelning/<% = sBild %>/<% = sBild %>_visa_info.asp?e=XXXX_GAMEID" title="XXXX_GAME">XXXX_CUTGAME</a>
        <% If canEdit Then %>
          <span style="float: right;">
            <img src="<% = config_GFXLocation %>icons/redigera.gif" alt="R" title="Redigera" title="" onclick="OpenCollection('<% = sEditWr %>',XXXX_GAMEID,XXXX_POSTID,'edit');">
            <img src="<% = config_GFXLocation %>icons/radera.gif" alt="X" title="Radera" onclick="DeleteCollection('<% = sEditWr %>',XXXX_POSTID);">
          </span>
        <% End If %>
        <div class="nf_collectionbar" style="background-image: url('<% = config_GFXLocation %>icons/samling/samling_alla_<% = sBild %>.png');">
          <img alt="" title="Box" src="<% = config_GFXLocation %>icons/samling/noXXXX_CBOX.png">
          <img alt="" title="Media" src="<% = config_GFXLocation %>icons/samling/noXXXX_CMEDIA.png">
          <img alt="" title="Manual" src="<% = config_GFXLocation %>icons/samling/noXXXX_CMANUAL.png"> 
          <img alt="" title="Extra" src="<% = config_GFXLocation %>icons/samling/noXXXX_CEXTRA.png"> 
        </div>
      </div>
      <% ' #### /SAMLINGEN #### %>
    </div>
  
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_minibox nf_blue">
        <h4>Hela samlingen</h4>
        <div class="nf_inside">
          <p><span style="float: right;"><strong><% = text_AntalSpel %></strong></span> <img src="<% = config_GFXLocation %>icons/spel.png"> <a href="?list=spel&amp;m=<% = lMedlem %>">Spel</a> </p>
          <p><span style="float: right;"><strong><% = text_AntalKonsoler %></strong></span> <img src="<% = config_GFXLocation %>icons/konsol.png"> <a href="?list=konsol&amp;m=<% = lMedlem %>">Konsoler</a> </p>
          <p><span style="float: right;"><strong><% = text_AntalTillbehor %></strong></span> <img src="<% = config_GFXLocation %>icons/tillbehor.png"> <a href="?list=tillbehor&amp;m=<% = lMedlem %>">Tillbehör</a> </p>
        </div>
      </div>
      
      <% If sList = "SPEL" Then %>
        <div class="nf_minibox nf_blue">
          <h4>Spel per konsol</h4>
          <div class="nf_inside">
            <% For zx = 1 To lstKonsol(0) %>
              <% lData = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Spel LEFT JOIN cms_Spel ON biSpelID = cms_Spel.sID WHERE sKonsol = " & CLng(zx) & " AND biAnv = " & CLng(anvID))(0) %>
              <p><span style="float: right;"><strong><% = lDATA %></strong></span> <img src="<% = config_GFXLocation %>icons/konsol/<% = zx %>.png"> <a href="?list=spel&amp;k=<% = zx %>&amp;m=<% = lMedlem %>"><% = lstKonsol(zx) %></a> </p>
            <% Next %>
          </div>
        </div>
      <% End If %>

    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->