<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  sQ           = Trim(MakeLegal(GetQ("q", "ABC", 255)))
  
  If Len(sQ) > config_MinSearch Then
    ' #### FIX TEXT STR�NG ####
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
  
    sQUrl = Server.URLEncode(sQ) 
    sQ    = sEncode(sQ)
    
    filter_All = "&amp;q=" & sQUrl

    dataSQL = "SELECT TOP 250 aaID, aaDatumPublicerad, aaTitel, aaKategori, aaStatus, aaAnvandarArt, aaText, aaFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 2 AND cBindID = aaID) " & _
              "FROM cms_Artiklar " & _
              "LEFT JOIN CONTAINSTABLE(cms_Artiklar, *, '" & p & "') AS ct ON aaID = ct.[KEY] " & _
              "WHERE Rank > 0 AND aaStatus = 4 AND aaDatumPublicerad <= '" & Now & "' " & _
              "ORDER BY Rank DESC, aaTitel ASC"
              
    noMsg   = "Inga tr�ffar p� [<strong>" & sEncode(q) & "</strong>], prova att bredda din s�kning."
    aSearch = True
  Else
    dataSQL = "SELECT aaID, aaDatumPublicerad, aaTitel, aaKategori, aaStatus, aaAnvandarArt, aaText, aaFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 2 AND cBindID = aaID) " & _
              "FROM cms_Artiklar " & _
              "WHERE aaStatus = 4 AND aaDatumPublicerad <= '" & Now & "' " & _
              "ORDER BY aaDatumPublicerad DESC"
              
    noMsg   = "Det finns inga artiklar att visa."
    aSearch = False
  End If

  RS_Open 1, dataSQL, False
  
    If rsDB(1).EOF Then
      any_Art = False
    Else
      any_Art = True
      list_Art = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  ' ### Kommentarer
  If Not config_LockDown_Kommentarer Then
    RS_Open 1, "SELECT TOP 10 cID, cTextM, fsBB_Anv.aAnvNamn, cDatum, cAvdelning, cBindID, cms_Nyheter.nTitel, cms_Recensioner.rTitel, cms_Artiklar.aaTitel FROM cms_Kommentarer " & _
               "LEFT JOIN fsBB_Anv ON cms_Kommentarer.cAnv = fsBB_Anv.aID " & _
               "LEFT JOIN cms_Nyheter ON cms_Kommentarer.cBindID = cms_Nyheter.nID " & _
               "LEFT JOIN cms_Recensioner ON cms_Kommentarer.cBindID = cms_Recensioner.rID " & _
               "LEFT JOIN cms_Artiklar ON cms_Kommentarer.cBindID = cms_Artiklar.aaID " & _
               "WHERE cAvdelning = 2 " & _
               "ORDER BY cDatum DESC", False
    
      If rsDB(1).EOF Then
        any_Kommentarer = False
      Else
        any_Kommentarer = True
        list_Kommentarer = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  
  If any_Art Then
    CreatePaging 50, UBound(list_Art, 2)
    CreatePagingChooser
  End If
  
  If pagingOnPage < 1 Then pagingOnPage = 1
%>

<%
  ' ## Globala variabler ##
  If aSearch Then
    page_Title    = sEncode(q) & " - S�k - Sida " & pagingOnPage & " - Artiklar"
  Else
    page_Title    = "Artiklar - Sida " & pagingOnPage
  End If
  
  page_Header   = "Nintendo artiklar"
  page_WhereAmI = "&gt; <a href='default.asp' title='G� till &quot;Artiklar&quot; ...'>Artiklar</a> "
  page_SelMenu  = "texter"
  page_Slide    = "artiklar"
  
  page_description    = "Arkiv �ver alla v�ra artiklar p� N-Forum.se, Nintendo Forum. Du �r p� sida " & pagingOnPage & "."
  page_keywords       = "artiklar, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/artiklar/">Artiklar</a></span>Nintendo Artiklar</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <div class="nf_msg nf_largesearch">
        <p><strong>S�k efter artiklar...</strong></p>
        <form>
          <input style="width: 564px;" type="text" maxlength=255 name="q" value="<% = sQ %>"> 
          <input style="float: right; width: 80px; font-weight: bold;" type="submit" value="S�k">
        </form>
      </div>
    
      <% If any_Art Then %>
          
          <ul class="nf_list">
            <% If aSearch Then %><li class="nf_listsplit"> S�ktr�ffar </li><% End If %>
          
            <%
              remPost = CDate("2050-01-01 00:00:00")
              For zx = pagingBOF To pagingEOF
                If zx > UBound(list_Art, 2) Then Exit For
                
                If Not aSearch Then
                  If DateDiff("m", list_Art(1, zx), remPost) <> 0 OR DateDiff("yyyy", list_Art(1, zx), remPost) <> 0 Then 
                    remPost = list_Art(1, zx)
                    %>
                    <li class="nf_listsplit"> <% = MonthName(Month(list_Art(1, zx))) %>&nbsp;<% = Year(list_Art(1, zx)) %> </li>
                    <%
                  End If
                End If
                %>
                  <li>
                    <% If CLng(list_Art(7, zx)) > 0 Then %>
                      <div class="nf_front" style="height: 120px; background-image: url('<% = config_ImageLocation %>?e=<% = list_Art(7, zx) %>&w=180&h=120')"><p>&nbsp;</p></div>
                    <% Else %>
                      <div class="nf_front nf_front_art"><p><% If list_Art(5, zx) Then %>Anv�ndarartikel<% Else %>Artikel<% End If %></p></div>
                    <% End If %>
                    <div class="nf_data">
                      <h3><a href="artikel_visa.asp?e=<% = list_Art(0, zx) %>" title="<% = sEncode(list_Art(2, zx)) %>"><% = sEncode(list_Art(2, zx)) %></a></h3>
                      <span class="nf_medium nf_gray nf_bold"><% If list_Art(5, zx) Then %>Anv�ndarartikel<% Else %>Artikel<% End If %> / <% = lstKonsol(list_Art(3, zx)) %> / <% = DatumReplace(list_Art(1, zx)) %></span>
                      <p style="line-height: 18px;"><% = sEncode(CutText(BBCode_Remove(list_Art(6, zx)),180)) %></p>
                      
                      <div class="nf_morebtn">
                        <a href="artikel_visa.asp?e=<% = list_Art(0, zx) %>">L�s mer ...</a>
                        <a href="artikel_visa.asp?e=<% = list_Art(0, zx) %>#kommentarer" <% If CLng(list_Art(8, zx)) > 0 Then %>class="nf_hint"<% End If %>><% If CLng(list_Art(8, zx)) = 0 Then %>Kommentera texten<% ElseIf CLng(list_Art(8, zx)) = 1 Then %>1 kommentar<% Else %><% = list_Art(8, zx) %> kommentarer<% End If %></a>
                      </div>
                    </div>
                  </li>
                <%
              Next
            %>
          </ul>
          
          <div class="nf_paging">
            <a href="default.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">��</a> |
            
              <% For Each zx In pagingPages %>
                <% If zx = "..." Then %>
                  ... |
                <% Else %>
                  <a href="default.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
                <% End If %>
              <% Next %>
              
            | <a href="default.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">��</a>
          </div>
      <% Else %>
        <div class="nf_msg"><p><% = noMsg %></p></div>
      <% End If %>
    </div>
      
    <div class="nf_datablock nf_size_onethird">
    
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <% If CONST_LOGIN Then %>
        <div class="nf_minibox nf_green">
          <h4>Skapa artikel</h4>
          <div class="nf_inside">
            <p class="nf_huge nf_center"><strong><a href="ny_artikel.asp" title="">Ny artikel</a></strong></p>
          </div>
        </div>
      <% End If %>
    
      <div class="nf_minibox nf_blue">
        <h4>Artiklar p� N-Forum.se</h4>
        <div class="nf_inside">
          <p>H�r finner du alla artiklar som finns p� N-Forum.se</p>
        </div>
      </div>
      
      <div class="nf_minibox nf_green">
        <h4>S�ktips</h4>
        <div class="nf_inside">
          <p>S�k p� minst <strong>3</strong> tecken.</p>
          <p>Om du f�r f�r m�nga resultat prova att vara mer specifik och anv�nd fler ord.</p>
          <p>Du kan <strong>INTE</strong> anv�nda termer s� som AND, OR och liknande</p>
        </div>
      </div>
      
      <% ' #### KOMMENTARER %>
      <% If any_Kommentarer Then %>
        <div class="nf_minibox nf_blue">
          <h4>Senaste kommentarerna</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_Kommentarer, 2) %>
                <%
                Select Case CLng(list_Kommentarer(4, zx))
                  Case 0 : sTypText = "Nyhet"     : sPrefixT = "news" : sXText = "nyheten"      : sCatText = "nyheter/nyheter"          : sMinCat = lstKategori(list_Kommentarer(3, zx))
                  Case 1 : sTypText = "Recension" : sPrefixT = "rec"  : sXText = "recensionen"  : sCatText = "recensioner/recension"    : sMinCat = lstKonsol(list_Kommentarer(3, zx))
                  Case 2 : sTypText = "Artikel"   : sPrefixT = "art"  : sXText = "artikeln"     : sCatText = "artiklar/artikel"         : sMinCat = lstKonsol(list_Kommentarer(3, zx))
                End Select
                %>
              
                <li onclick="location.href='/avdelning/<% = sCatText %>_visa.asp?e=<% = list_Kommentarer(5, zx) %>#kommentar_<% = list_Kommentarer(0, zx) %>';">
                  <a href="/avdelning/<% = sCatText %>_visa.asp?e=<% = list_Kommentarer(5, zx) %>#kommentar_<% = list_Kommentarer(0, zx) %>"><% = sTypText %> / <% = sEncode(CutText(list_Kommentarer(6+list_Kommentarer(4, zx), zx),28)) %></a>
                  <span><% = DatumReplace(list_Kommentarer(3, zx)) %> av <% = sEncode(CutText(list_Kommentarer(2,zx),20)) %></span>
                  <p style="width: 241px; padding: 0; margin: 0; font-weight: normal;"><% = sEncode(CutText(BBCode_Remove(list_Kommentarer(1, zx)),80)) %></p>
                </li>
              <% Next %>
            </ul>
          </div>
        </div>
      <% End If %>
        
    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->