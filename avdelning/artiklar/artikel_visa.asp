<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  ' ## Hämta all data ##
  lID = GetQ("e", "123", 0)
  RS_Open 1, "SELECT * FROM cms_Artiklar LEFT JOIN fsBB_Anv ON cms_Artiklar.aaSkapadAv = fsBB_Anv.aID WHERE aaDatumPublicerad <= '" & Now & "' And aaStatus = 4 AND aaID = " & lID, False
  
    If rsDB(1).EOF Then Response.Redirect("default.asp")
    
    text_Titel      = rsDB(1)("aaTitel")
    text_Text       = sEncode(rsDB(1)("aaText"))
    text_TextE      = sEncode(CutText(BBCode_Remove(rsDB(1)("aaText")),200))
    text_AvNamn     = sEncode(rsDB(1)("aNamn"))
    text_AvIn       = sEncode(rsDB(1)("aAnvNamn"))
    text_AvID       = rsDB(1)("aID")
    text_Kategori   = rsDB(1)("aaKategori")
    text_Publicerad = rsDB(1)("aaDatumPublicerad")
    text_AnvArt     = rsDB(1)("aaAnvandarArt")
  
  RS_Close 1
  
  ' ### Bilder
  RS_Open 1, "SELECT bID, baArtikel, baBildText, baBild, baID FROM cms_Bind_Artikel_Img LEFT JOIN cms_Bild ON cms_Bind_Artikel_Img.baBild = cms_Bild.bID WHERE baArtikel = " & lID & " ORDER BY baID ASC", False
  
    If rsDB(1).EOF Then
      any_Images = False
    Else
      any_Images = True
      list_Images = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  ' ### Fler artiklar
  RS_Open 1, "SELECT TOP 10 aaID, aaTitel, aaDatumPublicerad, aaKategori FROM cms_Artiklar WHERE aaDatumPublicerad <= '" & Now & "' AND aaStatus = 4 ORDER BY aaDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_XArt     = False
    Else
      any_XArt     = True
      list_XArt    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler foruminlägg
  If Not config_LockDown_Forum Then
    ' #### FIX TEXT STRÄNG ####
      q = LCase(Trim(text_Titel))
      
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
  
    RS_Open 1, "SELECT TOP 8 tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, fIcon, fName, Rank " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN CONTAINSTABLE(fsBB_Tradar, tTextM, '" & p & "') AS ct ON tbTrad.tID = ct.[KEY] " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "WHERE Rank > 0 AND tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tStatus_Raderad = 0" & _
               "ORDER BY Rank DESC, tAmne ASC", False
    
      If rsDB(1).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(1).GetRows(8)
      End If
    
    RS_Close 1
  End If
  
  ' ## Kommentarer
  lAvdID = 2
  RS_Open 1, "SELECT cID, cTextM, cAnv, cDatum, " & _
             "fsBB_Anv.aAnvNamn, fsBB_Anv.aID, fsBB_Anv.aAvatar, fsBB_Anv.aTimeStamp, fsBB_Anv.aAktiveraPM " & _
             "FROM cms_Kommentarer " & _
             "LEFT JOIN fsBB_Anv ON cms_Kommentarer.cAnv = aID " & _
             "WHERE cAvdelning = " & CLng(lAvdID) & " AND cBindID = " & CLng(lID) & " " & _
             "ORDER BY cDatum ASC", False
   
     If rsDB(1).EOF Then
      any_Comments = False
    Else
      any_Comments = True
      list_Comments = rsDB(1).GetRows
    End If
   
  RS_Close 1
%>

<%
  ' ## Globala variabler ##
  page_Title    = sEncode(text_Titel) & " - Artiklar"
  page_Header   = sEncode(text_Titel)
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Artiklar&quot; ...'>Artiklar</a> " & _ 
                  "&gt; <a href='artikel_visa.asp?e=" & lID & "' title='Gå till &quot;" & sEncode(text_Titel) & "&quot; ...'>" & sEncode(text_Titel) & "</a> " & _
                  "&gt; Artikel"
  page_SelMenu  = "texter"
  page_Slide    = "artiklar"
  
  page_description    = "Du visar just nu artikeln (" & sEncode(text_Titel) & ") i vår artikelavdelning på N-Forum.se, Nintendo Forum. " & Replace(text_TextE, vbCrlf, " ") & "..."
  page_keywords       = "visa artikel, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/artiklar/">Artiklar</a></span><% = sEncode(text_Titel) %></h1>
      <h4><% = lstKonsol(text_Kategori) %> / Publicerad <% = DatumReplace(text_Publicerad) %> av <a href="/avdelning/medlem/?m=<% = text_AvIn %>"><% = text_AvIn %></a> <% If text_AnvArt Then %><span style="color: #060;"> / Användarartikel</span><% End If %></h4>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <div class="nf_text">
        <p><% = BBCode(text_Text, False) %></p>
      </div>
      
      <div class="nf_sharebox"><!--#INCLUDE FILE="../../__INC/_sharecode.asp"--></div>
      
      <!-- ### KOMMENTARER ### -->
      
        <ul class="nf_list">
          <li class="nf_listsplit"> Kommentarer <a name="kommentarer"> </a> </li>
          <% If any_Comments Then %>
            <% For zx = 0 To UBound(list_Comments, 2) %>
              <% If bOdd Then bOdd = False Else bOdd = True %>
              <li class="nf_comments <% If Not bOdd Then Response.Write("nf_odd") %>" style="padding: 20px;"> 
                <span class="nf_comments_date"><a href="#kommentar_<% = list_Comments(0, zx) %>"><% = DatumReplace(list_Comments(3, zx)) %></a> <% If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then %> | <a href="javascript: doActionWithPrompt('/_action/deletecomment.asp?e=<% = list_Comments(0, zx) %>','Vill du ta bort kommentaren?');">Radera</a><% End If %></span>
                <span class="nf_comments_name"><a href="/avdelning/medlem/?m=<% = list_Comments(4, zx) %>"><% = list_Comments(4, zx) %></a> <a name="kommentar_<% = list_Comments(0, zx) %>"> </a></span>

                <p style="width: 100%;"><% = TinyCode(list_Comments(1, zx)) %></p>
              </li>
            <% Next %>
          <% End If %>
        </ul>
        
        <% If Not any_Comments Then %><div class="nf_msg"><p> Det finns inga kommentarer, bli först att kommentera. </p></div><% End If %>
        
        <% If CONST_LOGIN Then %>
          <form method="POST" action="/_action/postcomment.asp">
            <div class="nf_form">
              <p><strong>Kommentera</strong></p>
            
              <div class="nf_falt"><textarea name="aMsg" style="height: 250px; width: 576px"></textarea></div>
              
              <div class="nf_falt nf_buttons">
                <input type="hidden" name="avd" value="<% = lAvdID %>">
                <input type="hidden" name="e" value="<% = lID %>">
                <input type="submit" style="font-weight: bold;" value="Posta">
              </div>
    
            </div>
          </form>
        <% Else %>
          <div class="nf_msg nf_green">
            <p style="text-align: center;"><em>Du måste <strong><a href="<% = config_NotLoggedIn %>">logga in</a></strong> för att kunna lämna kommentarer.</em></p>
            <p style="text-align: center;"><em>Om du inte redan har en användare kan du <strong><a href="/avdelning/medlem/registreradig.asp">bli medlem</a> GRATIS!</strong>.</em></p>
          </div>
        <% End If %>
      
      <!-- ### /KOMMENTARER ### -->
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <% If any_Images Then %>
        <div class="nf_minibox nf_blue">
          <h4>Bilder</h4>
          <div class="nf_inside nf_imgbox">
            <% For zx = 0 To UBound(list_Images, 2) %>
              <a href="<% = config_ImageLocation %>?e=<% = list_Images(0, zx) %>&amp;w=800&amp;h=600" rel="lightbox[bilder]" title="<% = sEncode(list_Images(2, zx)) %>" target="_blank"><img src="<% = config_ImageLocation %>?e=<% = list_Images(0, zx) %>&amp;w=80&amp;h=80" title="<% = sEncode(list_Images(2, zx)) %>" alt="<% = sEncode(list_Images(2, zx)) %>"></a>
            <% Next %>
          </div>
        </div>
      <% End If %>

      
      <% If any_XArt Then %>
        <div class="nf_minibox nf_blue">
          <h4>Artiklar</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XArt, 2) %>
                <li onclick="location.href='/avdelning/artiklar/artikel_visa.asp?e=<% = list_XArt(0, zx) %>';"><a href="/avdelning/artiklar/artikel_visa.asp?e=<% = list_XArt(0, zx) %>" title="<% = sEncode(list_XArt(1, zx)) %>"><% = sEncode(CutText(list_XArt(1, zx), 32)) %></a><% = lstKonsol(list_XArt(3, zx)) %> / <% = DatumReplace(list_XArt(2, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/artiklar/">Visa alla artiklar</a></p>
          </div>
        </div>
      <% End If %>
      
      <!-- ## FORUMINLÄGG ## -->
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
            <p><a href="/avdelning/forum/nyainlagg.asp">Visa alla foruminlägg</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /FORUMINLÄGG ## -->
      
    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->