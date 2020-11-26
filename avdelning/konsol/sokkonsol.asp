<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0

  sQ           = Trim(MakeLegal(GetQ("q", "ABC", 255)))
  text_Konsol  = GetQ("konsol", "123", 0)
  If text_Konsol < 0 Then text_Konsol = 0
  
  If text_Konsol > 0 Then 
    sSokIForum = "AND kKonsol = " & CLng(text_Konsol)
  Else
    sSokIForum = ""
  End if

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
    
    RS_Open 1, "SELECT kID, tTitel, kKonsol, tSortNo, tBoxart_BoxFram, tBoxart_Manual, tBoxart_Konsol, tRegion, tRelease, tID, " & _ 
               "(SELECT COUNT(biID) FROM cms_Bind_Anv_Konsol WHERE biTitelID = cms_KonsolTitlar.tID AND biAnv = " & CLng(lAnvID) & ") AS tListadAntal " & _
               "FROM cms_KonsolTitlar " & _ 
               "LEFT JOIN CONTAINSTABLE(cms_KonsolTitlar, *, '" & p & "') AS ct ON tID = ct.[KEY] " & _
               "LEFT JOIN cms_Konsol ON cms_KonsolTitlar.tKonsolID = cms_Konsol.kID " & _ 
               "WHERE Rank > 0 AND kSynlig = 1 " & sSokIForum & " " & _
               "ORDER BY Rank DESC, tTitel ASC", False
    
      If rsDB(1).EOF Then
        any_Consoles = False
        sMess = "Inga tr�ffar p� [<strong>" & sEncode(q) & "</strong>], prova att bredda din s�kning."
      Else
        any_Consoles = True
        list_Consoles = rsDB(1).GetRows
      End If
    
    RS_Close 1
  Else
    If Len(sQ) = 0 Then
      sMess = "Du har inte gjort n�gon s�kning �nnu."
    Else
      sMess = "Du m�ste s�ka p� minst <strong>tre (3)</strong> tecken."
    End If
    any_Games = False
  End If
  
  sQUrl = Server.URLEncode(sQ)
  sQ    = sEncode(sQ)
  
  filter_All = "&amp;konsol=" & text_Konsol & "&amp;q=" & sQUrl
%>

<%
  ' ## Globala variabler ##
  If any_Games Then
    page_Title    = "[" & sQ & "] - S�k - Konsol"
  Else
    page_Title    = "S�k - Konsol"
  End If
  
  page_Header   = "S�k konsol"
  page_WhereAmI = "&gt; S�k konsol "
  page_SelMenu  = "databas"
  page_Slide    = "konsoler"
  
  page_description  = "S�k efter nintendo konsoler p� N-Forum.se, Nintendo Forum."
  page_keywords     = "s�k konsol, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/konsol/">Konsoler</a></span>S�k konsoler</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <div class="nf_msg">
        <form>
        
          <select name="konsol" style="width: 569px;">
            <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> Alla konsoler </option>
            <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
            <% For zx = 1 To lstKonsol(0) %>
              <option value=<% = zx %> style="padding: 1px 0 1px 10px;" <% If CLng(text_Konsol) = CLng(zx) Then Response.Write(" selected") %>> <% = lstKonsol(zx) %> </option>
            <% Next %>
          </select> 
        
          <input style="width: 564px;" type="text" maxlength=255 name="q" value="<% = sQ %>"> 
          <input style="float: right; width: 80px; font-weight: bold;" type="submit" value="S�k">
        </form>
      </div>
      
      <% If any_Consoles Then %>
        <% CreatePaging 50, UBound(list_Consoles, 2) %>
        <% CreatePagingChooser %>
        
        <div class="nf_msg">
          <p>Du visar just nu s�ktr�ff <strong><% = pagingBOF+1 %></strong>-<strong><% = pagingEOF+1 %></strong> av <strong><% = pagingNumOfPosts %></strong> och �r p� sidan <strong><% = pagingOnPage %></strong> av <strong><% = pagingNumOfPages %></strong>.</p>
        </div>
        
        <ul class="nf_list">
          <li class="nf_listsplit"> S�ktr�ffar </li>
          <%
            For zx = pagingBOF To pagingEOF
              If zx > UBound(list_Consoles, 2) Then Exit For
              
              miniBox = 0
              If CLng(list_Consoles(5, zx)) > 0 Then miniBox = list_Consoles(5, zx)
              If CLng(list_Consoles(6, zx)) > 0 Then miniBox = list_Consoles(6, zx)
              If CLng(list_Consoles(4, zx)) > 0 Then miniBox = list_Consoles(4, zx)
              %>
          
                <li>
                  <div class="nf_tiny">
                    <% If CLng(miniBox) > 0 Then %>
                      <img src="<% = config_ImageLocation %>?e=<% = miniBox %>&amp;w=50&amp;h=50&amp;err=no">
                    <% Else %>
                      <img src="<% = config_GFXLocation %>img/noimg_24x24.gif">
                    <% End If %>
                  </div>
                  <div class="nf_data">
                    <h4>
                      <img src="<% = config_GFXLocation %>icons/flags/<% = CLng(list_Consoles(7, zx)) %>.png" alt="<% = lstRegion(CLng(list_Consoles(7, zx))) %>" title="Region: <% = lstRegion(CLng(list_Consoles(7, zx))) %>">
                      <a href="konsol_visa_info.asp?e=<% = list_Consoles(9, zx) %>" title="<% = sEncode(list_Consoles(1, zx)) %>"><% = sEncode(list_Consoles(1, zx)) %></a>
                    </h4>
                    <span class="nf_medium nf_gray nf_bold"><% = lstKonsol(list_Consoles(2, zx)) %></span>
                  </div>
                  <div class="nf_extend">
                    <% If CONST_LOGIN Then %>
                      <img src="<% = config_GFXLocation %>icons/plus_lrg.png" style="float: right; cursor: pointer;" alt="+" title="L�gg till titeln i din samling." onclick="OpenCollection('console',<% = list_Consoles(9, zx) %>,0,'list')">
                    <% Else %>
                      <img src="<% = config_GFXLocation %>icons/plus_lrg_bw.png" style="float: right;" alt="+" title="Du m�ste vara inloggad f�r att kunna lista dina konsoler.">
                    <% End If %>
                    <% If CONST_LOGIN Then %><img src="<% = config_GFXLocation %>icons/listed.gif" style="display: <% If CLng(list_Consoles(10, zx)) = 0 Then Response.Write("none") Else Response.Write("block") %>;" id="listicon_<% = list_Consoles(9, zx) %>" alt="LIST" title="Titeln finns i din samling."><% End If %>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
        
        <div class="nf_paging">
          <a href="sokkonsol.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">��</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="sokkonsol.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
          | <a href="sokkonsol.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">��</a>
        </div>
      <% Else %>
        <div class="nf_msg nf_red">
          <p><% = sMess %></p>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
    
      <div class="nf_minibox nf_blue">
        <h4>S�ktips</h4>
        <div class="nf_inside">
          <p>S�k p� minst <strong>3</strong> tecken.</p>
          <p>Om du f�r f�r m�nga resultat prova att vara mer specifik och anv�nd fler ord.</p>
          <p>Du kan <strong>INTE</strong> anv�nda termer s� som AND, OR och liknande</p>
        </div>
      </div>
      
    </div>

  </div>
  
<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->