<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  ' ## H�mta all data ##
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0
  
  lID = GetQ("e", "123", 0)
  RS_Open 1, "SELECT *, " & _
             "(SELECT COUNT(biID) FROM cms_Bind_Anv_Tillbehor WHERE biTitelID = cms_TillbehorTitlar.tID AND biAnv = " & CLng(lAnvID) & ") AS tListadAntal " & _
             "FROM cms_TillbehorTitlar " & _
             "LEFT JOIN cms_Tillbehor ON cms_TillbehorTitlar.tTillbehorID = cms_Tillbehor.iID " & _
             "WHERE tID = " & CLng(lID), False
  
    If rsDB(1).EOF Then Response.Redirect("default.asp")
    
    text_ID         = CLng(rsDB(1)("tTillbehorID"))
    
    text_Titel      = sEncode(rsDB(1)("tTitel"))
    text_TitelRaw   = rsDB(1)("tTitel")
    text_Region     = FixNum(rsDB(1)("tRegion"))
    text_RegKod     = sEncode(rsDB(1)("tRegionsKod"))
    text_Release    = sEncode(rsDB(1)("tRelease"))
    text_Konsol     = lstKonsol(rsDB(1)("iKonsol"))
    text_KonsolID   = rsDB(1)("iKonsol")
    text_TillbehorID= FixNum(rsDB(1)("tTillbehorID"))
    
    text_ListadAntal  = CLng(rsDB(1)("tListadAntal"))
    
    text_LargeText  = Trim(BBCode(sEncode(rsDB(1)("iTextM")), False))
      
    text_Img1       = CLng(rsDB(1)("tBoxart_BoxFram"))
    text_Img2       = CLng(rsDB(1)("tBoxart_BoxBak"))
    text_Img3       = CLng(rsDB(1)("tBoxart_Manual"))
    text_Img4       = CLng(rsDB(1)("tBoxart_Tillbehor"))
    
    If text_Img4 > 0 Then text_UseArt = text_Img4 : text_UseText = "Tillbeh�r"
    If text_Img3 > 0 Then text_UseArt = text_Img3 : text_UseText = "Manual"
    If text_Img2 > 0 Then text_UseArt = text_Img2 : text_UseText = "Boxart - Baksida"
    If text_Img1 > 0 Then text_UseArt = text_Img1 : text_UseText = "Boxart - Framsida"
  
  RS_Close 1
  
  RS_Open 1, "SELECT tID, tRegion, tTitel, tRelease FROM cms_TillbehorTitlar LEFT JOIN cms_Tillbehor ON cms_TillbehorTitlar.tTillbehorID = cms_Tillbehor.iID WHERE tTillbehorID = " & CLng(text_ID) & " ORDER BY tRelease ASC", False
  
    If rsDB(1).EOF Then
      any_Titles = False
    Else
      any_Titles = True
      list_Titles = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  ' #### H�MTA TITLAR I SAMLINGEN ####
  If CONST_LOGIN Then
    RS_Open 1, "SELECT biID, biTitelID, biBox, biManual, biTillbehor, biExtra, biInPris, tTitel, tRegion FROM cms_Bind_Anv_Tillbehor LEFT JOIN cms_Tillbehortitlar ON cms_Bind_Anv_Tillbehor.biTitelID = cms_Tillbehortitlar.tID WHERE biAnv = " & CONST_USERID & " AND biTillbehorID = " & CLng(text_ID), False
    
      If rsDB(1).EOF Then
        any_Samling = False
      Else
        any_Samling = True
        list_Samling  = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  ' ######################
  
  ' #### H�MTA ALLA SOM SAMLAR ####
    RS_Open 1, "SELECT aAnvNamn, tRegion, biBox, biManual, biTillbehor, biExtra, biOvrigt FROM fsBB_Anv " & _
               "LEFT JOIN cms_Bind_Anv_Tillbehor ON cms_Bind_Anv_Tillbehor.biAnv = aID " & _
               "LEFT JOIN cms_TillbehorTitlar ON cms_TillbehorTitlar.tID = cms_Bind_Anv_Tillbehor.biTitelID " & _
               "WHERE biTitelID IN (SELECT tID FROM cms_TillbehorTitlar WHERE tTillbehorID = " & CLng(text_TillbehorID) & ") AND aBlockadTill < '" & Date & "' AND aAktiverad = 1 AND aID <> " & CLng(lAnvID) & " ORDER BY biBox DESC, biTillbehor DESC, biManual DESC, biExtra DESC, aAnvNamn ASC", False
               
      If rsDB(1).EOF Then
        any_SomHar = False
      Else
        any_SomHar = True
        list_SomHar = rsDB(1).GetRows
      End If
    
    RS_Close 1
  ' ######################
  
  ' ### Fler foruminl�gg
  If Not config_LockDown_Forum Then
    ' #### FIX TEXT STR�NG ####
      p = ""
      q = ""
      ww = ""
    
      q = LCase(Trim(text_TitelRaw))
      
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
  
    RS_Open 2, "SELECT TOP 8 tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, fIcon, fName, Rank " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN CONTAINSTABLE(fsBB_Tradar, tTextM, '" & p & "') AS ct ON tbTrad.tID = ct.[KEY] " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "WHERE Rank > 0 AND tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tStatus_Raderad = 0" & _
               "ORDER BY Rank DESC, tAmne ASC", False
    
      If rsDB(2).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(2).GetRows(8)
      End If
    
    RS_Close 2
  End If
  
  If CLng(Session.Value("LastAddonSeen")) <> CLng(text_TillbehorID) Then
    Con.ExeCute("UPDATE cms_Tillbehortitlar SET tVisningar = tVisningar + 1 WHERE tID = " & CLng(lID))
    Session.Value("LastAddonSeen") = CLng(text_TillbehorID)
  End If
%>

<%
  ' ## Globala variabler ##
  page_Title    = text_Titel & " - Information - Tillbeh�r"
  page_Header   = text_Titel
  page_WhereAmI = "&gt; <a href='default.asp' title='G� till &quot;Tillbeh�r&quot; ...'>Tillbeh�r</a> " & _
                  "&gt; <a href='tillbehor_visa_info.asp?e=" & lID & "'>" & text_Titel & "</a> " & _
                  "&gt; Information"
  page_SelMenu  = "databas"
  page_Slide    = "tillbehor"
  
  page_description  = "Visar " & text_Titel & " till " & text_Konsol & " utgiven i " & GetRegion(text_Region) & " p� N-Forum.se, Nintendo Forum."
  page_keywords     = text_Titel & ", "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/tillbehor/">Tillbeh�r</a></span><img src="<% = config_GFXLocation %>icons/flags/<% = text_Region %>.png" alt="" title=""> <% = text_Titel %></h1>
      <h4><a href="default.asp?k=<% = text_KonsolID %>"><% = text_Konsol %></a></h4>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <% If any_Titles Then %>
        <div class="nf_msg">
          <p><strong>Finns i f�ljande utg�vor...</strong></p>
          <ul class="nf_rowlist">
            <% For zx = 0 To UBound(list_Titles, 2) %>
              <li onclick="location.href='tillbehor_visa_info.asp?e=<% = list_Titles(0, zx) %>'" <% If CLng(list_Titles(0, zx)) = CLng(lID) Then Response.Write(" class='c'") %>>
                <img src="<% = config_GFXLocation %>icons/flags/<% = list_Titles(1, zx) %>.png" alt="" title="">
                <a href="tillbehor_visa_info.asp?e=<% = list_Titles(0, zx) %>" title="<% = sEncode(list_Titles(2, zx)) %>"><% = sEncode(CutText(list_Titles(2, zx), 65)) %></a>
                <span style="float: right;"><% = sEncode(list_Titles(3, zx)) %></span>
              </li>
            <% Next %>
          </ul>
        </div>
      <% End If %>
      
      <% If text_UseArt > 0 Then %>
        <div class="nf_images">
          <div class="boxartblocker"></div>
          
          <% If text_Img1 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img1 %>&amp;w=640&amp;h=480" title="Boxart - Framsida" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img1 %>&amp;w=100&amp;h=100" title="Boxart - Framsida" alt="Boxart - Framsida">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <% If text_Img2 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img2 %>&amp;w=640&amp;h=480" title="Boxart - Baksida" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img2 %>&amp;w=100&amp;h=100" title="Boxart - Baksida" alt="Boxart - Baksida">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <% If text_Img3 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img3 %>&amp;w=640&amp;h=480" title="Manual" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img3 %>&amp;w=100&amp;h=100" title="Manual" alt="Manual">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <% If text_Img4 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img4 %>&amp;w=640&amp;h=480" title="Kassett/Media" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img4 %>&amp;w=100&amp;h=100" title="Tilbeh�r" alt="Konsol">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <div class="boxartblocker"></div>
        </div>
      <% End If %>
      
      <div class="nf_text">
        <p><strong>Information om <% = text_Titel %></strong></p>
        <% If Len(Trim(text_LargeText)) > 0 Then %>
          <p><% = text_LargeText %></p>
        <% Else %>
          <p><% = text_Titel %> �r ett tillbeh�r som tillh�r konsolfamiljen <strong><a href="default.asp?k=<% = text_KonsolID %>"><% = text_Konsol %></a></strong>. Denna konsol �r utgiven <strong><% = text_Release %></strong>.</p>
          <p>Vi har tyv�rr ingen mer information om detta tillbeh�r.</p>
        <% End If %>
      </div>
      
      <% If any_Images Then %>
        <div class="nf_images">
          <p><strong>Bilder...</strong></p>
          <% For zx = 0 To UBound(list_Images, 2) %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = list_Images(0, zx) %>&amp;w=640&amp;h=480" rel="lightbox[bilder]" title="<% = sEncode(list_Images(2, zx)) %>" target="_blank"><% End If %>
              <img src="<% = config_ImageLocation %>?e=<% = list_Images(0, zx) %>&amp;w=80&amp;h=80" title="<% = sEncode(list_Images(2, zx)) %>" alt="<% = sEncode(list_Images(2, zx)) %>">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% Next %>
        </div>
      <% End If %>
      
      <% If CONST_LOGIN Then %>
        <% ' #### SAMLINGEN #### %>
        <div class="nf_msg nf_green">
          <p><strong>Du har listat f�ljande kopior av tillbeh�ret...</strong></p>
          <ul class="nf_rowlist" id="titleListed_List" style="<% If Not any_Samling Then Response.Write("display: none;") %>"></ul>   
          <p class="nf_pretend_rowlist" id="titleListed_Mess" style="<% If any_Samling Then Response.Write("display: none;") %>">Du har inte listat tillbeh�ret.</p>
          <p style="text-align: center;"><input style="float: none;" type="button" onclick="OpenCollection('addon',<% = lID %>,0,'new');" value="L�gg till i samlingen"></p>
          <% If any_SomHar Then %>
            <p><strong><img id="toggleBt" src="<% = config_GFXLocation %>icons/plus.gif" onclick="toggleBox('listadav','toggleBt');" style="float: left; cursor: pointer; margin: 0 5px 0 0;"> <span style="float: left; margin: 1px 0 0 0;">Listat av f�ljande medlemmar...</span> </strong></p>
            <ul class="nf_rowlist" id="listadav" style="display: none;">
              <% For zx = 0 To UBound(list_SomHar, 2) %>
                <li>
                  <img src="<% = config_GFXLocation %>icons/flags/<% = list_SomHar(1,zx) %>.png" alt="" title="">
                  <a href="/avdelning/medlem/?m=<% = sEncode(list_SomHar(0,zx)) %>"><% = sEncode(list_SomHar(0,zx)) %></a>
                  <div class="nf_collectionbar" style="background-image: url('<% = config_GFXLocation %>icons/samling/samling_alla_tillbehor.png');">
                    <img alt="" title="Box" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(2,zx) Then Response.Write("blank") %>.png">
                    <img alt="" title="Tillbeh�r" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(4,zx) Then Response.Write("blank") %>.png">
                    <img alt="" title="Manual" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(3,zx) Then Response.Write("blank") %>.png"> 
                    <img alt="" title="Extra" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(5,zx) Then Response.Write("blank") %>.png"> 
                  </div>
                  <% If Len(list_SomHar(6,zx)) > 1 Then %><span style="float: right; margin: 0 5px 0 5px;" title="<% = sEncode(list_SomHar(6,zx)) %>"><% = sEncode(CutText(list_SomHar(6,zx), 50)) %></span> <% End If %>
                </li>
              <% Next %>
            </ul>
          <% Else %>
            <p><strong>Inga andra medlemmar har listat titeln. </strong></p>
          <% End If %>            
        </div>
        
        <div id="titleListed_Clone" style="display: none;">
          <img src="<% = config_GFXLocation %>icons/flags/XXXX_REGION.png" alt="" title="">
          <a href="spel_visa_info.asp?e=XXXX_GAMEID" title="XXXX_GAME">XXXX_CUTGAME</a>
          <span style="float: right;">
            <img src="<% = config_GFXLocation %>icons/redigera.gif" alt="R" title="Redigera" title="" onclick="OpenCollection('addon',XXXX_GAMEID,XXXX_POSTID,'edit');">
            <img src="<% = config_GFXLocation %>icons/radera.gif" alt="X" title="Radera" onclick="DeleteCollection('addon',XXXX_POSTID);">
          </span>
          <div class="nf_collectionbar" style="background-image: url('<% = config_GFXLocation %>icons/samling/samling_alla_tillbehor.png');">
            <img alt="" title="Box" src="<% = config_GFXLocation %>icons/samling/noXXXX_CBOX.png">
            <img alt="" title="Media" src="<% = config_GFXLocation %>icons/samling/noXXXX_CMEDIA.png">
            <img alt="" title="Manual" src="<% = config_GFXLocation %>icons/samling/noXXXX_CMANUAL.png"> 
            <img alt="" title="Extra" src="<% = config_GFXLocation %>icons/samling/noXXXX_CEXTRA.png"> 
          </div>
        </div>
        
        <script type="text/javascript">
          <% If any_Samling Then %>
            <% For zx = 0 To UBound(list_Samling, 2) %>
              <% titleTT = sEncode(CutText(list_Samling(7, zx), 65)) & "</a>" %>
              <% If list_Samling(2, zx) = True Then cBox = "blank" Else cBox = "" %>
              <% If list_Samling(4, zx) = True Then cMedia = "blank" Else cMedia = "" %>
              <% If list_Samling(3, zx) = True Then cManual = "blank" Else cManual = "" %>
              <% If list_Samling(5, zx) = True Then cExtra = "blank" Else cExtra = "" %>
              rh_cloneRow("titleListed_Clone", "titleListed_List", "titleListed_Row_", <% = list_Samling(0, zx) %>, "LI","REGION==<% = list_Samling(8, zx) %>;;GAMEID==<% = list_Samling(1, zx) %>;;GAME==<% = sEncode(list_Samling(7, zx)) %>;;CUTGAME==<% = titleTT %>;;POSTID==<% = list_Samling(0, zx) %>;;CBOX==<% = cBox %>;;CMEDIA==<% = cMedia %>;;CMANUAL==<% = cManual %>;;CEXTRA==<% = cExtra %>");
            <% Next %>
          <% End If %>
        </script>
        <% ' #### /SAMLINGEN #### %>
      <% Else %>
        <% ' #### SAMLINGEN #### %>
        <div class="nf_msg nf_green">
          <p style="text-align: center;">Du m�ste <strong><a href="/avdelning/medlem/loggain.asp">logga in</a></strong> f�r att kunna lista dina tillbeh�r.</p>
          <p style="text-align: center;">Om du inte redan har en anv�ndare kan du <strong><a href="/avdelning/medlem/registreradig.asp">bli medlem</a> GRATIS!</strong>.</p>
        </div>
        <% ' #### /SAMLINGEN #### %>
      <% End If %>
    </div>
    
    <div class="nf_datablock nf_size_onethird">
       
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <div class="nf_minibox nf_blue">
        <div class="nf_inside nf_boxart">
          <% If text_UseArt > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_UseArt %>&amp;w=800&amp;h=600" rel="lightbox" target="_blank" title="<% = text_UseText %>"><% End If %>
              <img src="<% = config_ImageLocation %>?e=<% = text_UseArt %>&amp;w=300&amp;h=300">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% Else %>
            <img src="<% = config_GFXLocation %>img/noimg_200x150.png">        
          <% End If %>
        </div>
      </div>
      
      <div class="nf_minibox">
        <h4>Dela med dig</h4>
        <div class="nf_inside">
          <!-- AddThis Button BEGIN -->
            <div class="addthis_toolbox addthis_default_style" addthis:title="<% = text_Titel %>" addthis:description="<% = text_LargeTextE %>">
              <a class="addthis_button_email" title="E-posta"></a>
              <a class="addthis_button_print" title="Skriv ut"></a>
              <span class="addthis_separator">|</span>
              <a class="addthis_button_facebook" title="Facebook"></a>
              <a class="addthis_button_twitter" title="Twitter"></a>
              <a class="addthis_button_digg" title="Digg"></a>
              <a class="addthis_button_pusha" title="Pusha"></a>
              <a class="addthis_button_blogger" title="Blogger"></a>
              <a class="addthis_button_delicious" title="Del.icio.us"></a>
              <a class="addthis_button_google" title="Google"></a>
            </div>
            <script type="text/javascript" src="http://s7.addthis.com/js/250/addthis_widget.js#username=nforum"></script>
          <!-- AddThis Button END -->
        </div>
      </div>
      
      <div class="nf_minibox nf_blue">
        <h4>Tillbeh�rsdata</h4>
        <div class="nf_inside">
          <div class="nf_rowhead">Release</div>
          <div class="nf_row"><% = text_Release %></div>
        </div>
      </div>
      
      <!-- ## FORUMINL�GG ## -->
      <% If any_Tradar Then %>
        <div class="nf_minibox nf_green">
          <h4>Forum</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_Tradar, 2) %>
                <%
                  isTheThread = False
                  If list_Tradar(4,zx) Then isTheThread = True
                %>
                <li onclick="location.href='/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>';"><a href="/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>" title="<% = sEncode(list_Tradar(1,zx)) %>"><% = sEncode(CutText(list_Tradar(1,zx), 32)) %></a><% = list_Tradar(8, zx) %> / <% = DatumReplace(list_Tradar(3,zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/forum/nyainlagg.asp">Visa alla foruminl�gg</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /FORUMINL�GG ## -->
      
    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->