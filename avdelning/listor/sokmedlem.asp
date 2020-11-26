<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  sQ           = Trim(MakeLegal(GetQ("q", "ABC", 255)))
  
  If Len(sQ) > config_MinSearch Then
  
    ' #### FIX TEXT STRÄNG ####
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
  
    RS_Open 1, "SELECT aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aInloggadSenast, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
               "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
               "WHERE aAnvNamn LIKE '%" & q & "%' AND aBlockadTill < '" & Date & "' AND aAktiverad = 1 ORDER BY aAnvNamn ASC", False
    
      If rsDB(1).EOF Then
        any_Reg = False
        sMess = "Inga träffar på [<strong>" & sEncode(q) & "</strong>], prova att bredda din sökning."
      Else
        any_Reg = True
        list_Reg = rsDB(1).GetRows
      End If
    
    RS_Close 1
  Else
    If Len(sQ) = 0 Then
      sMess = "Du har inte gjort någon sökning!"
    Else
      sMess = "Du har angett för få tecken!"
    End If
    any_Reg = False
  End If

  filter_all = "&amp;q=" & sQ  
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Sök medlem"
  page_Header   = "Sök medlemmar"
  page_WhereAmI = "&gt; <a href='senastreg.asp' title='Gå till &quot;Sök medleme&quot; ...'>Sök medlem</a> "
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
  page_description  = "Sök efter medlemmar på N-Forum.se, Nintendo Forum"
  page_keywords     = "sök medlem ,"
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Sök medlem</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      
      <div class="nf_msg">
        <form>
          <input style="width: 433px;" type="text" maxlength=255 name="q" value="<% = sQ %>"> 
          <input style="float: right; width: 80px; font-weight: bold;" type="submit" value="Sök">
        </form>
      </div>
      
      <% If any_Reg Then %>
        <% CreatePaging 50, UBound(list_Reg, 2) %>
        <% CreatePagingChooser %>
        
        <ul class="nf_list">
          <li class="nf_listsplit"> Sökträffar </li>
          <%
            For zx = pagingBOF To pagingEOF
              If zx > UBound(list_Reg, 2) Then Exit For
              %>
                <li>
                  <div class="nf_icon">
                    <% If list_Reg(5, zx) Then %>
                      <img src="<% = config_Avatar & "u" & Right("000000" & list_Reg(6, zx), 6) & ".jpg" %>" alt="Avatar">
                    <% End If %>
                  </div>
                  <div class="nf_data">
                    <h3><a href="/avdelning/medlem/?m=<% = list_Reg(0, zx) %>"><% = sEncode(list_Reg(0, zx)) %></a></h3>
                    <span class="nf_medium nf_gray"><% If Len(list_Reg(4, zx)) > 0 Then %><% = list_Reg(4, zx) %><% Else %><% = list_Reg(2, zx) %><% End If %></span> 
                    <span class="nf_bold">Inloggad senast: <% = DatumReplace(list_Reg(3, zx)) %></span>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
        
        <div class="nf_paging">
          <a href="sokmedlem.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">««</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="sokmedlem.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
          | <a href="sokmedlem.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">»»</a>
        </div>
      <% Else %>
        <div class="nf_msg">
          <p><% = sMess %></p>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
    
      <div class="nf_minibox nf_blue">
        <h4>Söktips</h4>
        <div class="nf_inside">
          <p>Sök på minst <strong>3</strong> tecken.</p>
          <p>Om du får för många resultat prova att vara mer specifik och använd fler ord.</p>
          <p>Du kan <strong>INTE</strong> använda termer så som AND, OR och liknande</p>
        </div>
      </div>
      
    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->