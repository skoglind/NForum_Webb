<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0

  RS_Open 1, "SELECT ksID, ksSkapadDatum, ksTitel, ksKategori1, ksStatus, ksKategori2, ksSkapadAv, ksTyp, ksSynlig, ksTextM " & _
             "FROM cms_KopSalj WHERE ksSkapadAv = " & CLng(lAnvID) & " " & _
             "ORDER BY ksSkapadDatum DESC", False
             
    If rsDB(1).EOF Then
      any_Ads = False
    Else
      any_Ads = True
      list_Ads = rsDB(1).GetRows
    End If
  
  RS_Close 1
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Mina annonser - Annonser"
  page_Header   = "Mina annonser"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Marknad&quot; ...'>Marknad</a>"
  page_SelMenu  = "buy"
  page_Slide    = "annonser"
  
  page_description    = "Alla dina egna annonser i vår köp- och sälj-avdelning på N-Forum.se, Nintendo Forum."
  page_keywords       = "mina annonser, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/annonser/">Annonser</a></span>Mina annonser</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      
        <% If any_Ads Then %>
          <% CreatePaging 50, UBound(list_Ads, 2) %>
          <% CreatePagingChooser %>
            
            <ul class="nf_list">
              <%
                For zx = pagingBOF To pagingEOF
                  If zx > UBound(list_Ads, 2) Then Exit For
                  %>
                    <li>
                      <div class="nf_header" style="background-color: <% = getColor(list_Ads(7, zx)) %>;">
                        <span><% = lstKSTyp(list_Ads(7, zx)) %></span>
                        <% If Not list_Ads(8, zx) Then %><div class="nf_xtra">Osynlig!</div><% End If %>
                        <% If list_Ads(4, zx) = 1 Then %><div class="nf_xtra">Avslutad!</div><% End If %>
                      </div>
                      <div class="nf_data">
                        <h3><a href="annons_visa.asp?e=<% = list_Ads(0, zx) %>" title="<% = sEncode(list_Ads(2, zx)) %>"><% = sEncode(list_Ads(2, zx)) %></a><a href="ny_annons.asp?e=<% = list_Ads(0, zx) %>"><img src="<% = config_GFXLocation %>icons/edit.png" style="float: right;" title="Redigera annonsen" alt="Redigera"></a></h3>
                        <span class="nf_medium nf_gray nf_bold"><% = lstKSKategori(list_Ads(3, zx)) %> / <% = DatumReplace(list_Ads(1, zx)) %></span>
                        <p><% = sEncode(CutText(BBCode_Remove(list_Ads(9, zx)),300)) %></p>
                        <span class="nf_small nf_bold">» <a href="annons_visa.asp?e=<% = list_Ads(0, zx) %>">Visa annonsen</a>...</span>
                      </div>
                    </li>
                  <%
                Next
              %>
            </ul>
            
            <div class="nf_paging">
              <a href="minaannonser.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">««</a> |
              
                <% For Each zx In pagingPages %>
                  <% If zx = "..." Then %>
                    ... |
                  <% Else %>
                    <a href="minaannonser.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
                  <% End If %>
                <% Next %>
                
              | <a href="minaannonser.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">»»</a>
            </div>
        <% Else %>
          <div class="nf_msg"><p>Det finns inga anonnser att visa.</p></div>
        <% End If %>
      </div>
      
      <div class="nf_datablock nf_size_onethird">
      
        <div class="nf_minibox nf_green">
          <h4>Skapa annons</h4>
          <div class="nf_inside">
            <p class="nf_huge nf_center"><strong><a href="ny_annons.asp" title="">Ny annons</a></strong></p>
          </div>
        </div>
      
        <div class="nf_minibox nf_blue">
          <h4>Dina annonser</h4>
          <div class="nf_inside">
            <p>Här finner du alla dina annonser som finns på N-Forum.se</p>
            <p>Annonserna syns bara i <strong><% = config_AdDays %></strong> dagar från att de läggs upp därefter får du spara om dem för att de ska synas på nytt.</p>
            <p>Kom ihåg att markera sålda annonser som sålda</p>
          </div>
        </div>
        
      </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->