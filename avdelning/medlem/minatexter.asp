<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  lMedlem = GetQ("m","ABC",50)
  If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME

  If Not dbUserExists(lMedlem) Then Response.Redirect("/")
  anvID = GetIDFromUsername(lMedlem)
  
  filter_list   = LCase(GetQ("list","ABC", 15))
  
  Select Case filter_list
    Case "recensioner"    : sList = "RECENSIONER"
    Case "artiklar"       : sList = "ARTIKLAR"
    Case Else             : sList = "HOME"
  End Select
  
  Select Case sList
    Case "RECENSIONER"
      ' #### RECENSIONER
      sAvdText    = "Recensioner"
      sAvdLank    = "/Avdelning/recensioner/recension_visa.asp?e="
      sTextAvd    = "rec"
      
      If CLng(anvID) <> CONST_USERID Then
        RS_Open 1, "SELECT rID, rDatumPublicerad, rTitel, rKategori, rStatus, rAnvandarRec, rText, rNotes FROM cms_Recensioner WHERE rDatumPublicerad <= '" & Now & "' And rStatus = 4 AND rSkapadAv = " & CLng(anvID) & " ORDER BY rDatumPublicerad DESC", False
      Else
        RS_Open 1, "SELECT rID, rDatumPublicerad, rTitel, rKategori, rStatus, rAnvandarRec, rText, rNotes FROM cms_Recensioner WHERE rStatus > 0 AND rSkapadAv = " & CLng(anvID) & " ORDER BY rStatus ASC, rDatumPublicerad DESC", False
      End If
                   
        If rsDB(1).EOF Then
          any_Text = False
        Else
          any_Text = True
          list_Text = rsDB(1).GetRows()
        End If
      
      RS_Close 1
      ' #### #### ####
    Case "ARTIKLAR"
      ' #### ARTIKLAR
      sAvdText    = "Artiklar"
      sAvdLank    = "/Avdelning/artiklar/artikel_visa.asp?e="
      sTextAvd    = "art"
      
      If CLng(anvID) <> CONST_USERID Then
        RS_Open 1, "SELECT aaID, aaDatumPublicerad, aaTitel, aaKategori, aaStatus, aaAnvandarArt, aaText, aaNotes FROM cms_Artiklar WHERE aaDatumPublicerad <= '" & Now & "' And aaStatus = 4 AND aaSkapadAv = " & CLng(anvID) & " ORDER BY aaDatumPublicerad DESC", False
      Else
        RS_Open 1, "SELECT aaID, aaDatumPublicerad, aaTitel, aaKategori, aaStatus, aaAnvandarArt, aaText, aaNotes FROM cms_Artiklar WHERE aaStatus > 0 AND aaSkapadAv = " & CLng(anvID) & " ORDER BY aaStatus ASC, aaDatumPublicerad DESC", False
      End If 
                   
        If rsDB(1).EOF Then
          any_Text = False
        Else
          any_Text = True
          list_Text = rsDB(1).GetRows()
        End If
      
      RS_Close 1
      ' #### #### ####
    Case Else
      sAvdText = lMedlem & "s texter"
  End Select
  
  If CLng(anvID) <> CONST_USERID Then
    text_AntalRecensioner      = Con.ExeCute("SELECT COUNT(rID) FROM cms_Recensioner WHERE rDatumPublicerad <= '" & Now & "' And rStatus = 4 AND rSkapadAv = " & CLng(anvID))(0)
    text_AntalArtiklar         = Con.ExeCute("SELECT COUNT(aaID) FROM cms_Artiklar WHERE aaDatumPublicerad <= '" & Now & "' And aaStatus = 4 AND aaSkapadAv = " & CLng(anvID))(0)
  Else
    text_AntalRecensioner      = Con.ExeCute("SELECT COUNT(rID) FROM cms_Recensioner WHERE rStatus > 0 AND rSkapadAv = " & CLng(anvID))(0)
    text_AntalArtiklar         = Con.ExeCute("SELECT COUNT(aaID) FROM cms_Artiklar WHERE aaStatus > 0 AND aaSkapadAv = " & CLng(anvID))(0)
  End If
  
  ' ### Fler recensioner
  RS_Open 1, "SELECT TOP 25 rID, rTitel, rDatumPublicerad, rKategori FROM cms_Recensioner WHERE rDatumPublicerad <= '" & Now & "' AND rStatus = 4 AND rSkapadAv = " & CLng(anvID) & " ORDER BY rDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_XRec     = False
    Else
      any_XRec     = True
      list_XRec    = rsDB(1).GetRows(25)
    End If
  
  RS_Close 1
  
  ' ### Fler artiklar
  RS_Open 1, "SELECT TOP 25 aaID, aaTitel, aaDatumPublicerad, aaKategori FROM cms_Artiklar WHERE aaDatumPublicerad <= '" & Now & "' AND aaStatus = 4 AND aaSkapadAv = " & CLng(anvID) & " ORDER BY aaDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_XArt     = False
    Else
      any_XArt     = True
      list_XArt    = rsDB(1).GetRows(25)
    End If
  
  RS_Close 1
  
  If CLng(anvID) <> CONST_USERID Then canEdit = False Else canEdit = True
  filter_All = "&amp;m=" & lMedlem & "&amp;list=" & LCase(sList)
%>
  
<%

  ' ## Globala variabler ##
  page_Title    = lMedlem & " - " & sAvdText & " - Medlem"
  page_Header   = lMedlem & "s " & Replace(LCase(sAvdText), "mina", "")
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> " & _
                  "&gt; Texter"
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
      <h1><% = sAvdText %></h1>
    </div>
  
    <% If sList <> "HOME" Then ' #### TEXTER %>
      <div class="nf_datablock nf_size_twothird">    
      <% If any_Text Then %>
        <% CreatePaging 50, UBound(list_Text, 2) %>
        <% CreatePagingChooser %>
  
        
        <div class="nf_msg">
          <p>Du visar just nu text <strong><% = pagingBOF+1 %></strong>-<strong><% = pagingEOF+1 %></strong> av <strong><% = pagingNumOfPosts %></strong> och är på sidan <strong><% = pagingOnPage %></strong> av <strong><% = pagingNumOfPages %></strong>.</p>
        </div>
        
        <ul class="nf_list">
          <%
            For zx = pagingBOF To pagingEOF
              If zx > UBound(list_Text, 2) Then Exit For
              %>
                <li> 
                  <div class="nf_icon">
                    <img src="<% = config_GFXLocation %>img/texter_<% = list_Text(4, zx) %>.png" alt="" title="">
                  </div>
                  <div class="nf_data">
                    <h3><a href="<% = sAvdLank %><% = list_Text(0, zx) %>" title="<% = sEncode(list_Text(2, zx)) %>"><% = sEncode(list_Text(2, zx)) %></a></h3>
                    
                    <% If list_Text(4, zx) = 4 Then %>
                      <span class="nf_medium nf_gray nf_bold">
                        <% If list_Text(1, zx) <= Now Then %>
                          Publicerad: <% = DatumReplace(list_Text(1, zx)) %>
                        <% Else %>
                          Kommer publiceras: <% = DatumReplace(list_Text(1, zx)) %>
                        <% End If %>
                      </span>
                    <% End If %>
                    
                    <% If list_Text(4, zx) = 1 Or list_Text(4, zx) = 3 Then %>
                      <p>
                        <a href="redigeratext.asp?e=<% = list_Text(0, zx) %>&amp;avd=<% = sTextAvd %><% = filterAll %>&amp;page=<% = pagingOnPage %>"><img src="<% = config_GFXLocation %>icons/edit.png" title="Redigera"></a>
                        <img src="<% = config_GFXLocation %>icons/del.png" title="Ta bort" onclick="doActionWithPrompt('_action/deletetext.asp?e=<% = list_Text(0, zx) %>&amp;avd=<% = sTextAvd %><% = filterAll %>&amp;page=<% = pagingOnPage %>','Vill du ta bort texten?');" style="cursor: pointer;">
                      </p>
                    <% Else %>
                      <p>
                        <img src="<% = config_GFXLocation %>icons/edit_gray.png" title="Redigera">
                        <img src="<% = config_GFXLocation %>icons/del_gray.png" title="Ta bort">
                      </p>
                    <% End If %>
                    
                    <span class="nf_small nf_bold">» <a href="<% = sAvdLank %><% = list_Text(0, zx) %>">Läs texten</a>...</span>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
        
        <div class="nf_paging">
          <a href="minatexter.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">««</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="minatexter.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
          | <a href="minatexter.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">»»</a>
        </div>
      <% Else %>
        <div class="nf_msg"><p>Det finns inga texter att visa.</p></div>
      <% End if %>
      </div>
    <% Else %>
        
      <!-- ### RECENSIONER ### -->
      <div class="nf_datablock nf_size_onethird">
        <div class="nf_msg nf_bigbutton">
          <a href="?list=recensioner&amp;m=<% = lMedlem %>" title="Visa alla recensioner">Recensioner</a>
        </div>
        
        <% If any_XRec Then %>
          <div class="nf_minibox nf_blue">
            <h4>Recensioner</h4>
            <div class="nf_inside nf_stylelist">
              <ul>
                <% For zx = 0 To UBound(list_XRec, 2) %>
                  <li onclick="location.href='/avdelning/recensioner/recension_visa.asp?e=<% = list_XRec(0, zx) %>';"><a href="/avdelning/recensioner/recension_visa.asp?e=<% = list_XRec(0, zx) %>" title="<% = sEncode(list_XRec(1, zx)) %>"><% = sEncode(CutText(list_XRec(1, zx), 32)) %></a><% = lstKonsol(list_XRec(3, zx)) %> / <% = DatumReplace(list_XRec(2, zx)) %></li>
                <% Next %>
              </ul>
              <p><a href="?list=recensioner&amp;m=<% = lMedlem %>">Visa alla recensioner</a></p>
            </div>
          </div>
        <% End If %>
      </div>
      
      <!-- ### ARTIKLAR ### -->
      <div class="nf_datablock nf_size_onethird">
        <div class="nf_msg nf_bigbutton">
          <a href="?list=artiklar&amp;m=<% = lMedlem %>" title="Visa alla artiklar">Artiklar</a>
        </div>
        
        <% If any_XArt Then %>
          <div class="nf_minibox nf_blue">
            <h4>Artiklar</h4>
            <div class="nf_inside nf_stylelist">
              <ul>
                <% For zx = 0 To UBound(list_XArt, 2) %>
                  <li onclick="location.href='/avdelning/artiklar/artikel_visa.asp?e=<% = list_XArt(0, zx) %>';"><a href="/avdelning/artiklar/artikel_visa.asp?e=<% = list_XArt(0, zx) %>" title="<% = sEncode(list_XArt(1, zx)) %>"><% = sEncode(CutText(list_XArt(1, zx), 32)) %></a><% = lstKonsol(list_XArt(3, zx)) %> / <% = DatumReplace(list_XArt(2, zx)) %></li>
                <% Next %>
              </ul>
              <p><a href="?list=artiklar&amp;m=<% = lMedlem %>">Visa alla artiklar</a></p>
            </div>
          </div>
        <% End If %>
      </div>

    <% End if %>
      
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_minibox">
        <h4>Texter</h4>
        <div class="nf_inside">
          <p><span style="float: right;"><strong><% = text_AntalRecensioner %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> <a href="?list=recensioner&amp;m=<% = lMedlem %>">Recensioner</a></p>
          <p><span style="float: right;"><strong><% = text_AntalArtiklar %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> <a href="?list=artiklar&amp;m=<% = lMedlem %>">Artiklar</a></p>
        </div>
      </div>

    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->