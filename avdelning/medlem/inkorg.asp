<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)

  RS_Open 1, "SELECT pID, pAmne, pLast, pBesvarad, pDatum, pFran, fsBB_Anv.aAnvNamn AS pAnvNamn FROM fsBB_PM " & _
             "LEFT JOIN fsBB_Anv ON pFran = fsBB_Anv.aID " & _
             "WHERE pTill = " & CLng(CONST_USERID) & " AND pRaderadTill = 0 ORDER BY pDatum DESC", False
  
    If rsDB(1).EOF Then
      any_PM = False
    Else
      any_PM = True
      list_PM = rsDB(1).GetRows
    End If
  
  RS_Close 1
%>
  
<%

  ' ## Globala variabler ##
  page_Title    = "Inkorg - PM - Medlem"
  page_Header   = "Inkorg"
  page_WhereAmI = "&gt; PM &gt; <a href='inkorg.asp' title='Gå till &quot;Inkorg&quot; ...'>Inkorg</a> "
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_u.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Inkorg (PM)</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <% If any_PM Then %>
        <% CreatePaging CONST_SET_PMSIDA, UBound(list_PM, 2) %>
        <% CreatePagingChooser %>
    
          
          <div class="nf_msg">
            <p>Du visar just nu PM <strong><% = pagingBOF+1 %></strong>-<strong><% = pagingEOF+1 %></strong> av <strong><% = pagingNumOfPosts %></strong> och är på sidan <strong><% = pagingOnPage %></strong> av <strong><% = pagingNumOfPages %></strong>.</p>
          </div>
          
          <ul class="nf_list">
            <%
              For zx = pagingBOF To pagingEOF
                If zx > UBound(list_PM, 2) Then Exit For
                %>
                  <li>
                    <% If list_PM(3, zx) Then %>
                      <div class="nf_mini" style="background-color: #f1e593;"><img src="<% = config_GFXLocation %>icons/pm_besvarad.gif"><p>Besvarad</p></div>
                    <% ElseIf list_PM(2, zx) Then %>
                      <div class="nf_mini"><img src="<% = config_GFXLocation %>icons/pm_last.gif"><p>Öppnad</p></div>
                    <% Else %>
                      <div class="nf_mini" style="background-color: #7de07a;"><img src="<% = config_GFXLocation %>icons/pm_olast.gif"><p>Oläst</p></div>
                    <% End If %>
                    
                    <div class="nf_data">
                      <h5><a href="pm_visa.asp?e=<% = list_PM(0, zx) %>" title="<% = sEncode(list_PM(1, zx)) %>"><% = sEncode(list_PM(1, zx)) %></a></h5>
                      <span class="nf_medium nf_gray nf_bold"><% = DatumReplace(list_PM(4, zx)) %> av <a href="/avdelning/medlem/?m=<% = list_PM(6, zx) %>"><% = list_PM(6, zx) %></a></span>
                    </div>
                  </li>
                <%
              Next
            %>
          </ul>
          
          <div class="nf_paging">
            <a href="inkorg.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">««</a> |
            
              <% For Each zx In pagingPages %>
                <% If zx = "..." Then %>
                  ... |
                <% Else %>
                  <a href="inkorg.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
                <% End If %>
              <% Next %>
              
            | <a href="inkorg.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">»»</a>
          </div>
      <% Else %>
        <div class="nf_msg"><p>Du har inga PM i din inkorg.</p></div>
      <% End If %>
      
    </div>
    
     <div class="nf_datablock nf_size_onethird">
        <div class="nf_minibox">
          <h4>PM - Personliga meddelanden</h4>
          <div class="nf_inside">
            <!--#INCLUDE FILE="_sidebar_pm.asp"-->
          </div>
        </div>
     </div>
      
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->